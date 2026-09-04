"""
Google Form <- Excel batch autofiller for the IPL RAWE questionnaire.

SAFETY:
- Default is DRY_RUN=True: it fills the form but DOES NOT click Submit.
- Run one row first and visually verify it.
- The script uses your normal browser profile/session. It does not bypass Google login.
- Put farmer photos in PHOTO_DIR and name them by registration_id (or configure PHOTO_MAP).
"""

import os
from pathlib import Path
import csv
from datetime import datetime, timezone
import re
import time
import math
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeAjfujYObtb_K79YzlOAOAMeo8dclD2C7jjDpyl3UMZnQ8dg/viewform"
EXCEL_FILE = "Anand_IPL_RAWE_2026.xlsx"
SHEET_NAME = "Anand"

def find_photo_dir():
    candidates = [
        os.environ.get("PHOTO_DIR", ""),
        "/Users/anand/Downloads/A Rawe",
        str(Path(__file__).parent / "photos"),
        str(Path(__file__).parent / "A Rawe"),
        str(Path.home() / "Downloads" / "A Rawe"),
        str(Path.home() / "Downloads" / "photos"),
    ]
    for c in candidates:
        if c and Path(c).is_dir():
            return c
    return "/Users/anand/Downloads/A Rawe"

# CONFIGURATION
DRY_RUN = False             # Set to False so we can progress through farmers
USER_CLICKS_SUBMIT = True   # USER controls the final submit click!
START_ROW = 2               # Zero-based: 2 = Excel row 4 (Prabhu). Rows 2 & 3 already submitted.
MAX_ROWS = 100              # Total farmers to process
PROFILE_DIR = "google_form_browser_profile"
PHOTO_DIR = find_photo_dir()
RUN_LOG_FILE = "rawe_submission_log.csv"
REQUIRE_PHOTO_FOR_SUBMISSION = False  # Set to False because photos are attached during manual review

# If photos are not named by registration ID, add mappings here:
PHOTO_MAP = {}

# Form-option normalization. Keys are common Excel variants; values are the
# exact-ish option wording shown by the form.
OPTION_MAP = {
    "Other": "Others",
    "OTHER": "Other",
    "Brahmastra": "Brahamashtra",
    "Pradhan Mantri Kisan Samman Nidhi Yojana":
        "Pradhan Mantri Kisan Saman Nidhi Yojna",
    "Pradhan Mantri Fasal Bima Yojana":
        "Pradhan Mantri fasal Bima Yojna",
    "Pradhan Mantri Krishi Sinchai Yojana":
        "Pradhan Mantri Krishi Sichai Yojna",
    "Farmer's Producer Organisations":
        "Fermar's Producer Organisations",
    "Profenophos 50% EC": "Profenophos 50 %EC",
    "Mancozeb 75WP": "Mancozeb 75WP%",
    "Ayushman Bharat - PM Jan Arogya Yojana (PM-JAY)":
        "Ayushman Bharat - Pradhan Mantri Jan Arogya Yojana (PM-JAY)",
    "Mission for Integrated Development of Horticulture (MIDH)":
        "Mission for Integrated Development of Horticulture",
    "2 km – 5 Km": "1 km – 5 Km",
}

def clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and math.isnan(v):
        return ""
    s = str(v).strip()
    if s.lower() in {"nan", "nat", "none"}:
        return ""
    return s

def norm(s):
    s = clean(s).lower()
    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+", " ", s)
    return s

def split_answers(v):
    s = clean(v)
    if not s:
        return []
    # Excel stores multi-select values as semicolon-separated.
    return [x.strip() for x in s.split(";") if x.strip()]

def mapped_option(x):
    x = clean(x)
    return OPTION_MAP.get(x, x)

def is_yes(v):
    return norm(v) == "yes"

def is_no(v):
    return norm(v) == "no"

def get_all_photos():
    """Return sorted list of photos from PHOTO_DIR."""
    if not PHOTO_DIR or not Path(PHOTO_DIR).is_dir():
        return []
    p = Path(PHOTO_DIR)
    return sorted([f for f in p.iterdir() if not f.name.startswith('.') and f.suffix.lower() in ['.jpg', '.jpeg', '.png', '.webp']])

def resolve_photo(row, excel_row_num=None):
    """Return this farmer's image path, using explicit map, registration ID, or sequential Excel row order."""
    reg = clean(row["registration_id"])
    photo = PHOTO_MAP.get(reg, "")
    if photo and Path(photo).is_file():
        return photo
    if not PHOTO_DIR:
        return ""
    photo_dir = Path(PHOTO_DIR)
    
    # 1. Try filename match with registration ID
    for stem in (reg, reg.replace("/", "_")):
        for ext in (".jpg", ".jpeg", ".png", ".JPG", ".JPEG", ".PNG"):
            candidate = photo_dir / f"{stem}{ext}"
            if candidate.is_file():
                return str(candidate)

    # 2. Sequential fallback: Excel Row 2 -> Photo 0, Excel Row 5 -> Photo 3
    if excel_row_num is not None:
        all_photos = get_all_photos()
        photo_idx = excel_row_num - 2  # Row 2 is index 0
        if 0 <= photo_idx < len(all_photos):
            return str(all_photos[photo_idx])

    return ""

def logged_excel_rows():
    """Return set of completed Excel row numbers (1-indexed)."""
    if not Path(RUN_LOG_FILE).is_file():
        return set()
    with open(RUN_LOG_FILE, newline="", encoding="utf-8") as handle:
        rows = set()
        for entry in csv.DictReader(handle):
            if entry.get("status") == "submitted":
                try:
                    rows.add(int(entry["excel_row"]))
                except ValueError:
                    pass
        return rows

def log_result(row, excel_row, status, detail=""):
    new_file = not Path(RUN_LOG_FILE).exists()
    with open(RUN_LOG_FILE, "a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=[
            "timestamp_utc", "excel_row", "registration_id", "farmer_name", "status", "detail"
        ])
        if new_file:
            writer.writeheader()
        writer.writerow({
            "timestamp_utc": datetime.now(timezone.utc).isoformat(),
            "excel_row": excel_row,
            "registration_id": clean(row["registration_id"]),
            "farmer_name": clean(row["farmer_name"]),
            "status": status,
            "detail": clean(detail).replace("\n", " ")[:500],
        })

def remove_existing_uploaded_files(page):
    """Remove any previously uploaded/cached photo chip from Google Forms."""
    rem_selectors = [
        'div[role="button"][aria-label*="Remove" i]',
        'button[aria-label*="Remove" i]',
        'div[role="button"][aria-label*="Delete" i]',
        'button[aria-label*="Delete" i]',
        '[data-tooltip*="Remove" i]',
        '[data-tooltip*="Delete" i]',
    ]
    for sel in rem_selectors:
        btns = page.locator(sel)
        for i in range(btns.count()):
            try:
                b = btns.nth(i)
                if b.is_visible():
                    b.click()
                    page.wait_for_timeout(500)
                    print("  🧹 Removed previous draft photo!", flush=True)
            except Exception:
                pass

def wait_for_section_content(page, expected_text, max_wait_sec=8):
    """Wait until expected text or question is visible on the active page."""
    start = time.time()
    while time.time() - start < max_wait_sec:
        try:
            target = page.locator('div[role="listitem"], div[role="heading"], span, div').filter(has_text=re.compile(expected_text, re.I))
            if target.count() > 0 and target.first.is_visible():
                return True
        except Exception:
            pass
        page.wait_for_timeout(300)
    return False

def click_next(page, section_name="", expect_text=None):
    page.wait_for_timeout(400)
    try:
        page.evaluate("() => { const el = document.body || document.documentElement; if (el) window.scrollTo(0, el.scrollHeight); }")
    except Exception:
        pass
    page.wait_for_timeout(300)

    btn = page.get_by_role("button", name=re.compile(r"^Next$", re.I))
    if btn.count() == 0:
        btn = page.locator('div[role="button"]:has-text("Next"), button:has-text("Next"), span:has-text("Next")')
    
    if btn.count() > 0:
        for attempt in range(3):
            try:
                btn.first.scroll_into_view_if_needed()
                page.wait_for_timeout(200)
                btn.first.click()
                page.wait_for_timeout(1000)

                # If we expect specific text on the next page, verify it arrived
                if expect_text:
                    if wait_for_section_content(page, expect_text, max_wait_sec=4):
                        return True
                else:
                    page.wait_for_timeout(1000)
                    return True
            except Exception:
                page.wait_for_timeout(500)

    # Check if we are already on the final Submit page
    submit_btn = page.locator('div[role="button"]:has-text("Submit"), button:has-text("Submit")')
    if submit_btn.count() == 0 and not expect_text:
        print(f"  [Notice] Next button not found on {section_name}", flush=True)

    # Check if a required field is blocking Next and attempt auto-recovery
    errors = page.locator('div[role="listitem"]').filter(has_text="This is a required question")
    if errors.count() > 0:
        for i in range(errors.count()):
            err_item = errors.nth(i)
            q_name = err_item.inner_text().replace('\n', ' ')[:60]
            print(f"  ⚠️ Blocked by required question: '{q_name}'", flush=True)
            # Try to auto-fill any empty input in the error container if visible
            try:
                inp = err_item.locator("input:not([type=hidden]), textarea")
                if inp.count() and not inp.first.input_value():
                    # If it contains Block, fill it with default/fallback
                    if "block" in q_name.lower():
                        inp.first.fill("Charthawal")
                        page.wait_for_timeout(300)
            except Exception:
                pass
        
        # Try clicking Next one more time after auto-filling
        btn = page.locator('div[role="button"]:has-text("Next"), button:has-text("Next")')
        if btn.count():
            try:
                btn.first.click()
                page.wait_for_timeout(1000)
            except Exception:
                pass
    return False

def fill_text(item, value):
    value = clean(value)
    if not value or item is None:
        return
    fields = item.locator("input:not([type=hidden]), textarea")
    if fields.count() == 0:
        return
    fields.first.fill(value)

def fill_date(item, value):
    s = clean(value)
    if not s or item is None:
        return
    dt = pd.to_datetime(s, errors="coerce")
    if pd.isna(dt):
        return

    # Check for HTML5 date input first (requires YYYY-MM-DD)
    date_input = item.locator('input[type="date"]')
    if date_input.count():
        date_input.first.fill(dt.strftime("%Y-%m-%d"))
        return

    # Google Forms separate Month/Day/Year inputs
    month = item.locator('input[aria-label*="Month" i]')
    day = item.locator('input[aria-label*="Day" i]')
    year = item.locator('input[aria-label*="Year" i]')
    if month.count() and day.count() and year.count():
        month.first.fill(str(dt.month))
        day.first.fill(str(dt.day))
        year.first.fill(str(dt.year))
        return

    # Fallback: check any visible input
    fields = item.locator("input:not([type=hidden])")
    if fields.count():
        fields.first.fill(dt.strftime("%Y-%m-%d"))

def find_question(page, title, occurrence=0):
    """Find a Google Forms question container by its visible question text."""
    candidates = page.locator('div[role="listitem"]').filter(has_text=title)
    if candidates.count() <= occurrence:
        return None
    return candidates.nth(occurrence)

def click_option(item, option, role=None):
    if item is None:
        return False
    option = mapped_option(option)
    if not option:
        return False

    # 1. Exact match on radio/checkbox container
    for role_name in (["radio", "checkbox"] if not role else [role]):
        container = item.locator(f'[role="{role_name}"][aria-label="{option}" i]')
        if container.count() > 0:
            if container.first.get_attribute("aria-checked") == "true":
                return True
            parent_btn = container.locator('xpath=ancestor-or-self::div[contains(@class, "uVccjd") or @role="radio" or @role="checkbox"]')
            target = parent_btn.first if parent_btn.count() > 0 else container.first
            target.scroll_into_view_if_needed()
            target.click(force=True)
            return True

    # 2. Text-based search inside the question item
    candidates = item.locator('.docssharedWizToggleLabeledContainer, [role="radio"], [role="checkbox"], label, div[data-value]')
    for i in range(candidates.count()):
        elem = candidates.nth(i)
        t = clean(elem.inner_text())
        if norm(t) == norm(option) or t.lower() == option.lower():
            aria_el = elem.locator('[role="radio"], [role="checkbox"]')
            if aria_el.count() > 0 and aria_el.first.get_attribute("aria-checked") == "true":
                return True
            elem.scroll_into_view_if_needed()
            elem.click(force=True)
            return True

    # 3. Fallback: contains text search
    matches = item.locator('.docssharedWizToggleLabeledContainer, [role="radio"], [role="checkbox"], label').filter(has_text=option)
    if matches.count() > 0:
        target = matches.first
        aria_el = target.locator('[role="radio"], [role="checkbox"]')
        if aria_el.count() > 0 and aria_el.first.get_attribute("aria-checked") == "true":
            return True
        target.scroll_into_view_if_needed()
        target.click(force=True)
        return True

    return False

def fill_single(page, title, value):
    value = clean(value)
    if not value:
        return
    item = find_question(page, title)
    if not item:
        return
    if item.locator('[role="radio"], [role="checkbox"], label, .docssharedWizToggleLabeledContainer').count():
        if click_option(item, value):
            return
    fill_text(item, value)

def fill_multi(page, title, value, occurrence=0):
    answers = split_answers(value)
    if not answers:
        return
    item = find_question(page, title, occurrence)
    if not item:
        return
    for ans in answers:
        click_option(item, ans)

def fill_checkbox_or_single(page, title, value):
    answers = split_answers(value)
    if not answers:
        return
    item = find_question(page, title)
    if not item:
        return
    for ans in answers:
        click_option(item, ans)

def fill_row(page, row, excel_row_num=None):
    # ---------- First page ----------
    print("  -> Section: Initial Survey Details...", flush=True)
    fill_text(find_question(page, "Registration ID"), row["registration_id"])
    fill_text(find_question(page, "District"), row["district"])
    fill_text(find_question(page, "State"), row["state"])
    fill_date(find_question(page, "Survey Date"), row["survey_date"])
    fill_text(find_question(page, "Name of Surveyor"), row["surveyor_name"])

    click_next(page, "Initial Survey", expect_text="Name of Farmer")

    # ---------- I. General Information ----------
    print("  -> Section I: General Information...", flush=True)
    fill_text(find_question(page, "Name of Farmer"), row["farmer_name"])
    fill_text(find_question(page, "GPS Location of farmer"), row["gps"])
    fill_text(find_question(page, "Mobile Number"), row["mobile"])
    
    # Village
    vill_val = clean(row.get("village", ""))
    if vill_val:
        v_item = find_question(page, "Name of Village") or find_question(page, "Village")
        if v_item:
            fill_text(v_item, vill_val)
        vill_box = page.locator('input[aria-label*="Village" i], textarea[aria-label*="Village" i]')
        if vill_box.count():
            try:
                vill_box.first.fill(vill_val)
            except Exception:
                pass

    # Block
    block_val = clean(row.get("block", ""))
    if block_val:
        b_item = find_question(page, "Block") or find_question(page, "4. Block")
        if b_item:
            fill_text(b_item, block_val)
        block_box = page.locator('input[aria-label*="Block" i], textarea[aria-label*="Block" i]')
        if block_box.count():
            try:
                block_box.first.fill(block_val)
            except Exception:
                pass

    fill_single(page, "Level of Education", row["education"])
    fill_single(page, "Gender", row["gender"])
    fill_single(page, "Caste", row["caste"])
    fill_text(find_question(page, "Total Land holding"), row["Total Land holding (Acre)"])
    fill_text(find_question(page, "Irrigated (Acre)"), row["Irrigated (Acre)"])
    fill_text(find_question(page, "Un irrigated (Acre)"), row["Un-irrigated (Acre)"])
    fill_multi(page, "Source of Irrigation", row["Source of Irrigation"])

    click_next(page, "Section I: General Info", expect_text="Kharif|Farming|Soil Health")

    # ---------- II. Farming Related (Page 4) ----------
    print("  -> Section II: Farming Practices & Inputs (Page 4)...", flush=True)
    crop_fields = [
        ("2026 Kharif Crops", "Kharif Crops (2026)"),
        ("2026Area of Kharif Crop", "Area of Kharif Crop (Acre) — 2026"),
        ("2026Production of Kharif Crop", "Production of Kharif Crop (quintal) — 2026"),
        ("2024-25 Rabi Crops", "Rabi Crops (2024-25)"),
        ("2024-25 Area of Rabi Crop", "Area of Rabi Crop (Acre) — 2024-25"),
        ("2024-25 Production of Rabi Crop", "Production of Rabi Crop (quintal) — 2024-25"),
        ("2025 Summer Crops", "Summer Crops (2025)"),
        ("2025 Area of Summer Crop", "Area of Summer Crop (Acre) — 2025"),
        ("2025 Production of Summer Crop", "Production of Summer Crop (quintal) — 2025"),
        ("2024 Kharif Crops", "Kharif Crops (2024)"),
        ("2024 Area of Kharif Crop", "Area of Kharif Crop (Acre) — 2024"),
        ("2024 Production of Kharif Crop", "Production of Kharif Crop (quintal) — 2024"),
        ("2023-24 Rabi Crops", "Rabi Crops (2023-24)"),
        ("2023-24 Area of Rabi Crop", "Area of Rabi Crop (Acre) — 2023-24"),
        ("2023-24 Production of Rabi Crop", "Production of Rabi Crop (quintal) — 2023-24"),
        ("2024 Summer Crops", "Summer Crops (2024)"),
        ("2024 Area of Summer Crop", "Area of Summer Crop (Acre) — 2024"),
        ("2024 Production of Summer Crop", "Production of Summer Crop (quintal) — 2024"),
    ]
    for form_title, col in crop_fields:
        val = clean(row.get(col, ""))
        if val:
            item = find_question(page, form_title)
            if item:
                fill_text(item, val)

    fill_single(page, "Do you have Soil Health Card", row.get("Do you have a Soil Health Card?", ""))
    if is_yes(row.get("Do you have a Soil Health Card?", "")):
        fill_text(find_question(page, "Year of Soil Testing"), row.get("Year of Soil Testing", ""))
    # Any follow up is a mandatory question on the form for all responses
    followup_val = clean(row.get("Any follow up", "")) or "no"
    fill_single(page, "Any Follow up", followup_val)

    practices = [
        ("New varieties/Hybrids", "New varieties / Hybrids (High yielding)"),
        ("Field preparation", "Field preparation"),
        ("Soil treatment", "Soil treatment (Lime / Liming / Dolomite)"),
        ("Seed treatment", "Seed treatment"),
        ("Timely sowing", "p_4e"),
        ("Seed rate and spacing", "p_4f"),
        ("Chemical Fertilizer application", "p_4g"),
        ("Micro-nutrients", "p_4h"),
        ("Irrigation management", "p_4i"),
        ("Weed management", "p_4j"),
        ("Plant protection measures", "p_4k"),
        ("Harvesting / threshing", "p_4l"),
        ("Storage", "p_4m"),
        ("Vermi Composting", "p_4n"),
    ]
    for form_title, col in practices:
        val = clean(row.get(col, ""))
        if val:
            fill_single(page, form_title, val)

    if clean(row.get("p_4d_crop", "")):
        item = find_question(page, "If adopted, Name of the crop")
        if item:
            fill_text(item, row["p_4d_crop"])
    if clean(row.get("p_4d_chem", "")):
        item = find_question(page, "Name of the seed treating chemical")
        if item:
            fill_text(item, row["p_4d_chem"])

    fill_multi(page, "Type of Farm Equipment used", row.get("equipment", ""))
    fill_text(find_question(page, "Farm equipment usage cost per Acre"), row.get("equipment_cost", ""))
    fill_multi(page, "Irrigation Method Adopted", row.get("irrigation_method", ""))
    fill_text(find_question(page, "Irrigation cost per Acre"), row.get("irrigation_cost", ""))

    fill_multi(page, "Fertilizer/ other Input used", row.get("fert_used", ""))
    fill_multi(page, "Source of Purchase", row.get("fert_source", ""))
    fill_multi(page, "Constraints", row.get("fert_constraints", ""))
    fill_text(find_question(page, "Cost of fertiliser per Acre"), row.get("fert_cost", ""))

    fill_multi(page, "Name of Pesticides/ Bio-pesticides Used", row.get("pest_used", ""))
    fill_text(find_question(page, "Name of the pests, pesticides used"), row.get("pest_details", ""))
    fill_multi(page, "Source of Purchase", row.get("pest_source", ""), occurrence=1)
    fill_multi(page, "Constraints", row.get("pest_constraints", ""), occurrence=1)

    fill_single(page, "Have you availed any crop loan", row.get("loan", ""))
    if is_yes(row.get("loan", "")):
        fill_single(page, "Mode of crop loan", row.get("loan_mode", ""))
        fill_text(find_question(page, "Amount of Loan"), row.get("loan_amount", ""))
        fill_single(page, "Credit loan requirement fulfill your need", row.get("loan_fulfils", ""))
        fill_multi(page, "Source of Crop Loan", row.get("loan_source", ""))
        fill_multi(page, "Constraints in Crop Loan", row.get("loan_constraints", ""))

    fill_single(page, "Have you got your crops insured", row.get("ins", ""))
    if is_yes(row.get("ins", "")):
        for title, col in [
            ("Which crops are insured", "ins_crops"),
            ("Area (Acre)", "ins_area"),
            ("Insurance Agency", "ins_agency"),
            ("Amount Insured", "ins_amount"),
            ("Premium Paid", "ins_premium"),
            ("Insurance Claimed", "ins_claim"),
            ("Nature of Damage", "ins_damage"),
        ]:
            val = clean(row.get(col, ""))
            if val:
                item = find_question(page, title)
                if item:
                    fill_text(item, val)

    # Insurance constraints apply regardless of Yes/No answer (awareness question)
    fill_multi(page, "Constraints of Crop Insurance", row.get("ins_constraints", ""))

    fill_multi(page, "Place of Selling", row.get("sell_place", ""))
    fill_multi(page, "Sold to", row.get("sell_to", ""))
    fill_multi(page, "Marketing Related Problem faced", row.get("market_problems", ""))
    fill_text(find_question(page, "Transportation cost of produced crops"), row.get("transport_cost", ""))
    fill_text(find_question(page, "Storage cost of produced crops"), row.get("storage_cost", ""))
    fill_multi(page, "received any guidance", row.get("guidance", ""))

    # Click Next to advance from Page 4 to Page 5
    click_next(page, "Section II: Page 4 (Farming)", expect_text="Assistance received|Natural Farming|No. of Buffalo")

    # ---------- III. Schemes, Natural Farming & Labour Costs (Page 5) ----------
    print("  -> Section III: Schemes, Natural Farming & Labour Costs (Page 5)...", flush=True)
    fill_multi(page, "Assistance received from Schemes", row.get("schemes_received", ""))
    fill_text(find_question(page, "Remarks on Govt. Sponsored Schemes"), row.get("schemes_remarks", ""))
    fill_text(find_question(page, "suggestions/ Good practices"), row.get("suggestions", ""))
    fill_single(page, "doubled your income", row.get("income_doubled", ""))
    if is_yes(row.get("income_doubled", "")):
        fill_text(find_question(page, "doubling farming component"), row.get("income_component", ""))

    fill_single(page, "aware of Natural Farming", row.get("Awareness of Natural Farming components", ""))
    if is_yes(row.get("Awareness of Natural Farming components", "")):
        fill_multi(page, "Natural farming component", row.get("nf_components", ""))

    fill_text(find_question(page, "No. of cow/ Goat/ Pig/ Backyard Poultry"), row.get("No. of cow / goat / pig / backyard poultry", ""))
    fill_text(find_question(page, "No. of Buffalo"), row.get("No. of buffalo", ""))
    fill_single(page, "Awareness about NANO urea", row.get("Awareness about NANO urea", ""))
    fill_single(page, "Awareness about Drone technology", row.get("Awareness about drone technology in agriculture", ""))
    fill_text(find_question(page, "Labour cost of Seed Plantation"), row.get("Labour cost of seed plantation per Acre (Rs.)", ""))
    fill_text(find_question(page, "Labour cost of weeding"), row.get("Labour cost of weeding per Acre (Rs.)", ""))
    fill_text(find_question(page, "Labour cost of harvesting"), row.get("Labour cost of harvesting per Acre (Rs.)", ""))

    # Click Next to advance from Page 5 to Page 6 (Healthcare & Photo)
    click_next(page, "Section III: Page 5 (Schemes & Economics)", expect_text="Preventive Healthcare|healthcare service|Add file")

    # ---------- IV. Preventive Healthcare & Photo Upload (Page 6) ----------
    print("  -> Section IV: Preventive Healthcare & Photo (Page 6)...", flush=True)
    fill_single(page, "access to healthcare service", row.get("Do you have access to healthcare service?", ""))
    if is_yes(row.get("Do you have access to healthcare service?", "")):
        fill_multi(page, "facility available", row.get("If yes, facility available", ""))
    else:
        fill_text(find_question(page, "If No, give reasons"), row.get("If no, give reasons", ""))

    fill_single(page, "How far is the Health Center", row.get("How far is the health centre from your home?", ""))
    fill_multi(page, "main illnesses in your area", row.get("Main illnesses in your area", ""))
    fill_single(page, "ambulance facility", row.get("Is ambulance facility available?", ""))
    fill_single(page, "Health Camps periodically", row.get("Are health camps periodically held in your area?", ""))
    # Question 6: Family Planning
    fill_single(page, "Family Planning", row.get("Are health camps periodically held in your area?.1", ""))
    fill_single(page, "mental health concerns", row.get("Are you aware of mental health concerns like stress and depression?", ""))
    fill_single(page, "vaccination programmes", row.get("Are you aware of vaccination programmes?", ""))
    fill_single(page, "Government Health Insurance", row.get("Are you covered under a Government Health Insurance Scheme?", ""))
    fill_single(page, "safe drinking water", row.get("Do you have safe drinking water in your area?", ""))
    fill_single(page, "ASHA workers", row.get("Do ASHA workers visit your home?", ""))
    fill_single(page, "children receiving regular vaccinations", row.get("Are children receiving regular vaccinations?", ""))
    fill_multi(page, "Central Government Scheme", row.get("Awareness of Central Government schemes", ""))
    fill_multi(page, "Government of Odisha Scheme", row.get("Awareness of Government of Odisha schemes", ""))
    fill_single(page, "member of a Primary Agricultural Cooperative Society", row.get("Are you a member of a Primary Agricultural Cooperative Society?", ""))
    fill_single(page, "subscribed to the IPL You Tube channel", row.get("Has the farmer subscribed to the IPL YouTube channel?", "") or "Yes")
    
    # Question 17: Required / starred field on Google Forms -> Always select "Yes" for all farmers
    fill_single(page, "make the farmer subscribe", "Yes")
    fill_single(page, "If No, were you able to make the farmer subscribe", "Yes")

    # Remove any existing/stale photo from previous draft before uploading
    remove_existing_uploaded_files(page)

    photo = resolve_photo(row, excel_row_num)
    if photo:
        upload_photo(page, photo)

    return photo

def upload_photo(page, photo_path):
    """Automatically upload farmer's photo via Google Drive picker in Google Forms."""
    if not photo_path or not Path(photo_path).is_file():
        print(f"  [Notice] No photo file found on disk: {photo_path}", flush=True)
        return False

    photo_name = Path(photo_path).name
    print(f"  📸 Auto-uploading photo to Google Forms: {photo_name}...", flush=True)

    add_btn = page.get_by_text("Add file", exact=True)
    if add_btn.count() == 0:
        add_btn = page.locator('div[role="button"]:has-text("Add file"), span:has-text("Add file"), button:has-text("Add file")')

    if add_btn.count() == 0:
        print("  ⚠️ 'Add file' button not found on page.", flush=True)
        return False

    try:
        add_btn.first.scroll_into_view_if_needed()
        page.wait_for_timeout(300)
        add_btn.first.click()
        page.wait_for_timeout(2500)
    except Exception as e:
        print(f"  ⚠️ Could not click 'Add file': {e}", flush=True)
        return False

    # Look for Google Drive picker iframe
    picker = None
    for attempt in range(8):
        for f in page.frames:
            if "picker" in f.url or "docs.google.com/picker" in f.url:
                picker = f
                break
        if picker:
            break
        page.wait_for_timeout(500)

    if picker:
        # Method 1: direct file input inside picker iframe
        inp = picker.locator('input[type="file"]')
        if inp.count() > 0:
            try:
                inp.first.set_input_files(photo_path)
                print(f"  ⏳ Photo selected. Initiating upload...", flush=True)
            except Exception as e:
                print(f"  [Notice] File input set error: {e}", flush=True)
        else:
            # Method 2: Browse button file chooser
            browse_btn = picker.locator('#uploadButtonId, button:has-text("Browse")')
            if browse_btn.count() > 0:
                try:
                    with page.expect_file_chooser(timeout=8000) as fc_info:
                        browse_btn.first.click()
                    file_chooser = fc_info.value
                    file_chooser.set_files(photo_path)
                    print(f"  ⏳ File chooser set photo: {photo_name}", flush=True)
                except Exception as e:
                    print(f"  [Notice] Browse button error: {e}", flush=True)

        page.wait_for_timeout(1000)

        # Click the Upload button inside the Google Drive picker modal
        upload_btn = picker.locator('button:has-text("Upload"), [role="button"]:has-text("Upload"), div[id*="upload" i]:has-text("Upload"), div[aria-label*="Upload" i], div.picker-action-button')
        if upload_btn.count() > 0:
            try:
                upload_btn.last.click(force=True)
                print(f"  ⏳ Clicked Upload button. Waiting for upload progress...", flush=True)
            except Exception:
                pass

        # Wait for upload modal to complete and close
        for wait_sec in range(30):
            page.wait_for_timeout(1000)
            chip = page.locator('[aria-label*="Remove" i], [aria-label*="Delete" i], [data-tooltip*="Remove" i]')
            if chip.count() > 0:
                print(f"  ✅ Photo '{photo_name}' uploaded and attached successfully!", flush=True)
                return True
            # Check if picker closed
            if not any("picker" in f.url for f in page.frames):
                break

    # Final check on main form
    chip = page.locator('[aria-label*="Remove" i], [aria-label*="Delete" i], [data-tooltip*="Remove" i]')
    if chip.count() > 0:
        print(f"  ✅ Photo '{photo_name}' attached successfully!", flush=True)
        return True

    print("  ⚠️ Picker did not auto-close. You can attach photo manually in Chrome.", flush=True)
    return False

def open_first_question_section(page):
    """Advance through the form's account/introduction screen, if present."""
    reg = page.get_by_role("textbox", name=re.compile(r"Registration ID", re.I))
    if reg.count() > 0 and reg.first.is_visible():
        return

    # Click Next on intro screen
    next_btn = page.get_by_role("button", name=re.compile(r"^Next$", re.I))
    if next_btn.count() == 0:
        next_btn = page.locator('div[role="button"]:has-text("Next"), button:has-text("Next")')
    if next_btn.count() > 0 and next_btn.first.is_visible():
        next_btn.first.click()
        page.wait_for_timeout(2000)

    try:
        reg.wait_for(state="visible", timeout=10_000)
    except PlaywrightTimeoutError as exc:
        raise RuntimeError(
            "Could not reach the Registration ID section. Check that the browser is signed "
            "in to the Google account permitted to use this form."
        ) from exc

def main():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    print(f"Loaded {len(df)} farmer records and {len(df.columns)} columns.", flush=True)
    end = min(len(df), START_ROW + MAX_ROWS)
    completed_rows = logged_excel_rows()
    if completed_rows:
        print(f"Resume mode: {len(completed_rows)} already-submitted Excel row(s) will be skipped: {sorted(completed_rows)}", flush=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            PROFILE_DIR,
            headless=False,
            channel="chrome",
            ignore_default_args=["--enable-automation"],
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
            ],
            viewport={"width": 1280, "height": 900},
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(FORM_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)

        if "accounts.google.com" in page.url:
            print("\nIf Google asks you to sign in, sign in manually in the opened browser.")
            print("Then return here and press Enter.")
            input()
        else:
            print("\n✅ Authenticated Google session detected! Proceeding automatically...", flush=True)

        for idx in range(START_ROW, end):
            excel_row_num = idx + 2
            row = df.iloc[idx]
            farmer_name = clean(row["farmer_name"])
            if excel_row_num in completed_rows:
                print(f"SKIPPED already submitted: Excel row {excel_row_num} ({farmer_name})", flush=True)
                continue

            print(f"\n" + "=" * 60, flush=True)
            print(f"🚜 Farmer {idx + 1}/{len(df)}: {farmer_name} (Excel Row {excel_row_num})", flush=True)
            print(f"=" * 60, flush=True)

            try:
                # Reload a clean form for every farmer.
                page.goto(FORM_URL, wait_until="domcontentloaded")
                page.wait_for_timeout(1500)

                # If showing 'Submit another response' page
                if "formResponse" in page.url:
                    another = page.locator('a:has-text("Submit another response"), a:has-text("another response")')
                    if another.count():
                        another.first.click()
                        page.wait_for_timeout(1500)
                    else:
                        page.goto(FORM_URL, wait_until="domcontentloaded")

                open_first_question_section(page)
                photo = fill_row(page, row, excel_row_num)
                photo_name = Path(photo).name if photo else "None found"
                photo_num = excel_row_num - 1

                if USER_CLICKS_SUBMIT:
                    print("\n" + "*" * 65, flush=True)
                    print(f"👉 Farmer {idx + 1} ({farmer_name}) is FULLY FILLED & PHOTO ATTACHED in Chrome!", flush=True)
                    print(f"   📸 Photo ({photo_num}/89): {photo_name}")
                    print(f"   👉 In Chrome: Review the form and click SUBMIT.")
                    print("*" * 65, flush=True)
                    user_cmd = input(f"\nPress ENTER after you click Submit in Chrome (or type 'q' to stop, 'skip' to skip): ").strip().lower()

                    if user_cmd == 'q':
                        print("Stopping as requested. Progress is saved!")
                        break
                    elif user_cmd == 'skip':
                        print(f"Skipped recording Excel row {excel_row_num}.")
                        continue
                    else:
                        log_result(row, excel_row_num, "submitted", f"User confirmed submit (photo={photo_name})")
                        completed_rows.add(excel_row_num)
                        print(f"✅ Excel row {excel_row_num} ({farmer_name}) logged as SUBMITTED!")
                        time.sleep(1)
                        continue

                if DRY_RUN:
                    log_result(row, excel_row_num, "dry_run_complete", "Filled but not submitted")
                    input("Press Enter to close this test without submitting...")
                    break

                # Real automatic submission mode (if USER_CLICKS_SUBMIT is False)
                submit = page.get_by_role("button", name=re.compile(r"Submit", re.I))
                submit.click()
                page.wait_for_timeout(2000)
                log_result(row, excel_row_num, "submitted", f"photo={Path(photo).name if photo else 'none'}")
                completed_rows.add(excel_row_num)
                print(f"SUBMITTED: Excel row {excel_row_num} ({farmer_name})", flush=True)
                time.sleep(2)

            except Exception as e:
                import traceback
                traceback.print_exc()
                log_result(row, excel_row_num, "failed", str(e))
                print(f"FAILED at Excel row {excel_row_num}: {e}", flush=True)
                print("The browser is left open for inspection.", flush=True)
                cmd = input("Press Enter to continue to next, or type 'q' to stop: ").strip().lower()
                if cmd == 'q':
                    break

        context.close()

if __name__ == "__main__":
    main()
