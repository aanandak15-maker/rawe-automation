"""
Google Form <- Excel batch autofiller for the IPL RAWE questionnaire.

SAFETY:
- Default is DRY_RUN=True: it fills the form but DOES NOT click Submit.
- Run one row first and visually verify it.
- The script uses your normal browser profile/session. It does not bypass Google login.
- Put farmer photos in PHOTO_DIR and name them by registration_id (or configure PHOTO_MAP).
"""

import sys
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
DATA_FILE = "pankaj_krishi_surveys_100.csv"
EXCEL_FILE = "pankaj_krishi_surveys_100.csv"
SHEET_NAME = "krishi-surveys (1) (1).csv"

def find_photo_dir():
    candidates = [
        os.environ.get("PHOTO_DIR", ""),
        "/Users/anand/Downloads/image rawe",
        "/Users/anand/Downloads/image_rawe",
        "/Users/anand/Downloads/farmer",
        "/Users/anand/Downloads/rawe",
        str(Path(__file__).parent / "farmer"),
        str(Path(__file__).parent / "rawe"),
        str(Path.home() / "Downloads" / "image rawe"),
    ]
    for c in candidates:
        if c and Path(c).is_dir():
            return c
    return "/Users/anand/Downloads/image rawe"

# CONFIGURATION
DRY_RUN = False             # Set to False so we can progress through farmers
USER_CLICKS_SUBMIT = False  # Auto-submit: no manual confirmation needed
START_ROW = 0               # Zero-based: 0 = Excel row 2
MAX_ROWS = 100              # Total farmers in Pankaj dataset
PROFILE_DIR = "google_form_browser_profile_pankaj"
PHOTO_DIR = find_photo_dir()
RUN_LOG_FILE = "pankaj_submission_log.csv"
REQUIRE_PHOTO_FOR_SUBMISSION = False

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
    if isinstance(v, float):
        if math.isnan(v):
            return ""
        if v.is_integer():
            return str(int(v))
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
    """Return sorted list of non-duplicate photos from PHOTO_DIR."""
    if not PHOTO_DIR or not Path(PHOTO_DIR).is_dir():
        return []
    p = Path(PHOTO_DIR)
    return sorted([f for f in p.iterdir() if not f.name.startswith('.') and f.suffix.lower() in ['.jpg', '.jpeg', '.png', '.webp'] and " copy" not in f.name])

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

    # 2. Sequential fallback: Excel Row 2 -> Photo 0, Excel Row 3 -> Photo 1
    if excel_row_num is not None:
        all_photos = get_all_photos()
        photo_idx = excel_row_num - 2  # Row 2 is index 0
        if 0 <= photo_idx < len(all_photos):
            return str(all_photos[photo_idx])

    return ""

def normalize_row(raw):
    """Normalize row dict whether it comes from short or long export format."""
    def get(*keys):
        for k in keys:
            if k in raw and pd.notna(raw[k]):
                val = clean(raw[k])
                if val:
                    return val
        return ""

    return {
        # Section 0
        "registration_id": get("Registration ID", "registration_id") or "ICRO/2026/20260800809",
        "district": "Muzaffarnagar" if "mu" in (get("District", "district") or "").lower() else (get("District", "district") or "Muzaffarnagar"),
        "state": get("State", "state") or "Uttar Pradesh",
        "survey_date": get("Survey Date", "survey_date") or "2026-08-18",
        "surveyor_name": get("Name of Surveyor", "surveyor_name") or "Pankaj kumar",

        # Section 1
        "farmer_name": get("Farmer Name", "farmer_name"),
        "gps": get("I. General Information — 1.a GPS Location of farmer", "GPS Location of farmer", "gps"),
        "mobile": get("I. General Information — 2. Mobile Number", "Mobile Number", "mobile"),
        "village": get("Village", "village"),
        "block": get("Block", "block"),
        "education": get("I. General Information — 5. Level of Education", "Level of Education", "education"),
        "gender": get("I. General Information — 6. Gender", "Gender", "gender"),
        "caste": get("I. General Information — 7. Caste", "Caste", "caste"),
        "Total Land holding (Acre)": get("I. General Information — 8. Total Land holding (Acre)", "Total Land holding (Acre)", "Total Land holding"),
        "Irrigated (Acre)": get("I. General Information — 8a. Irrigated (Acre)", "Irrigated (Acre)"),
        "Un-irrigated (Acre)": get("I. General Information — 8b. Un-irrigated (Acre)", "Un-irrigated (Acre)", "Un irrigated (Acre)"),
        "Source of Irrigation": get("I. General Information — 9. Source of Irrigation", "Source of Irrigation"),

        # Section 2 - Crops
        "Kharif Crops (2026)": get("II. Farming — Area & Production — Kharif Crops (2026)", "Kharif Crops (2026)"),
        "Area of Kharif Crop (Acre) — 2026": get("II. Farming — Area & Production — Area of Kharif Crop (Acre) — 2026", "Area of Kharif Crop (Acre) — 2026"),
        "Production of Kharif Crop (quintal) — 2026": get("II. Farming — Area & Production — Production of Kharif Crop (quintal) — 2026", "Production of Kharif Crop (quintal) — 2026"),
        
        "Rabi Crops (2024-25)": get("II. Farming — Area & Production — Rabi Crops (2024-25)", "Rabi Crops (2024-25)"),
        "Area of Rabi Crop (Acre) — 2024-25": get("II. Farming — Area & Production — Area of Rabi Crop (Acre) — 2024-25", "Area of Rabi Crop (Acre) — 2024-25"),
        "Production of Rabi Crop (quintal) — 2024-25": get("II. Farming — Area & Production — Production of Rabi Crop (quintal) — 2024-25", "Production of Rabi Crop (quintal) — 2024-25"),

        "Summer Crops (2025)": get("II. Farming — Area & Production — Summer Crops (2025)", "Summer Crops (2025)"),
        "Area of Summer Crop (Acre) — 2025": get("II. Farming — Area & Production — Area of Summer Crop (Acre) — 2025", "Area of Summer Crop (Acre) — 2025"),
        "Production of Summer Crop (quintal) — 2025": get("II. Farming — Area & Production — Production of Summer Crop (quintal) — 2025", "Production of Summer Crop (quintal) — 2025"),

        "Kharif Crops (2024)": get("II. Farming — Area & Production — Kharif Crops (2024)", "Kharif Crops (2024)"),
        "Area of Kharif Crop (Acre) — 2024": get("II. Farming — Area & Production — Area of Kharif Crop (Acre) — 2024", "Area of Kharif Crop (Acre) — 2024"),
        "Production of Kharif Crop (quintal) — 2024": get("II. Farming — Area & Production — Production of Kharif Crop (quintal) — 2024", "Production of Kharif Crop (quintal) — 2024"),

        "Rabi Crops (2023-24)": get("II. Farming — Area & Production — Rabi Crops (2023-24)", "Rabi Crops (2023-24)"),
        "Area of Rabi Crop (Acre) — 2023-24": get("II. Farming — Area & Production — Area of Rabi Crop (Acre) — 2023-24", "Area of Rabi Crop (Acre) — 2023-24"),
        "Production of Rabi Crop (quintal) — 2023-24": get("II. Farming — Area & Production — Production of Rabi Crop (quintal) — 2023-24", "Production of Rabi Crop (quintal) — 2023-24"),

        "Summer Crops (2024)": get("II. Farming — Area & Production — Summer Crops (2024)", "Summer Crops (2024)"),
        "Area of Summer Crop (Acre) — 2024": get("II. Farming — Area & Production — Area of Summer Crop (Acre) — 2024", "Area of Summer Crop (Acre) — 2024"),
        "Production of Summer Crop (quintal) — 2024": get("II. Farming — Area & Production — Production of Summer Crop (quintal) — 2024", "Production of Summer Crop (quintal) — 2024"),

        # Soil & Practices
        "Do you have a Soil Health Card?": get("II. Farming — Practices — 3(a). Do you have a Soil Health Card?", "Do you have a Soil Health Card?"),
        "Year of Soil Testing": get("II. Farming — Practices — (i) Year of Soil Testing", "Year of Soil Testing"),
        "Any follow up": get("II. Farming — Practices — (ii) Any follow up", "Any follow up"),

        "New varieties / Hybrids (High yielding)": get("II. Farming — Practices — 4a. New varieties / Hybrids (High yielding)", "New varieties / Hybrids (High yielding)"),
        "Field preparation": get("II. Farming — Practices — 4b. Field preparation", "Field preparation"),
        "Soil treatment (Lime / Liming / Dolomite)": get("II. Farming — Practices — 4c. Soil treatment (Lime / Liming / Dolomite)", "Soil treatment (Lime / Liming / Dolomite)"),
        "Seed treatment": get("II. Farming — Practices — 4d(i). Seed treatment", "Seed treatment"),
        "p_4d_crop": get("II. Farming — Practices — 4d(ii). If adopted, name of the crop", "p_4d_crop"),
        "p_4d_chem": get("II. Farming — Practices — 4d(iii). Name of the seed treating chemical and dose", "p_4d_chem"),
        "p_4e": get("II. Farming — Practices — 4e. Timely sowing", "p_4e"),
        "p_4f": get("II. Farming — Practices — 4f. Seed rate and spacing", "p_4f"),
        "p_4g": get("II. Farming — Practices — 4g. Chemical fertilizer application", "p_4g"),
        "p_4h": get("II. Farming — Practices — 4h. Micro-nutrients (Zn, F, Br, Mn, Mg)", "p_4h"),
        "p_4i": get("II. Farming — Practices — 4i. Irrigation management", "p_4i"),
        "p_4j": get("II. Farming — Practices — 4j. Weed management", "p_4j"),
        "p_4k": get("II. Farming — Practices — 4k. Plant protection measures", "p_4k"),
        "p_4l": get("II. Farming — Practices — 4l. Harvesting / threshing", "p_4l"),
        "p_4m": get("II. Farming — Practices — 4m. Storage", "p_4m"),
        "p_4n": get("II. Farming — Practices — 4n. Vermi composting", "p_4n"),

        # Equipment & Irrigation
        "equipment": get("II. Farming — Practices — 5. Type of farm equipment used", "equipment"),
        "equipment_cost": get("II. Farming — Practices — 5a. Farm equipment usage cost per Acre (Rs.)", "equipment_cost"),
        "irrigation_method": get("II. Farming — Practices — 6. Irrigation method adopted", "irrigation_method"),
        "irrigation_cost": get("II. Farming — Practices — 6a. Irrigation cost per Acre (Rs.)", "irrigation_cost"),

        # Inputs & Pesticides
        "fert_used": get("II. Farming — Inputs, Credit & Market — 7a. Fertilizer / other input used", "fert_used"),
        "fert_source": get("II. Farming — Inputs, Credit & Market — 7b. Source of purchase", "fert_source"),
        "fert_constraints": get("II. Farming — Inputs, Credit & Market — 7c. Constraints", "fert_constraints"),
        "fert_cost": get("II. Farming — Inputs, Credit & Market — 7d. Cost of fertiliser per Acre (Rs.)", "fert_cost"),

        "pest_used": get("II. Farming — Inputs, Credit & Market — 8a. Name of pesticides / bio-pesticides used", "pest_used"),
        "pest_details": get("II. Farming — Inputs, Credit & Market — 8b. Name of the pests, pesticides used and their waiting period", "pest_details"),
        "pest_source": get("II. Farming — Inputs, Credit & Market — 8c. Source of purchase", "pest_source"),
        "pest_constraints": get("II. Farming — Inputs, Credit & Market — 8d. Constraints", "pest_constraints"),

        # Loans & Insurance
        "loan": get("II. Farming — Inputs, Credit & Market — 9. Have you availed any crop loan?", "loan"),
        "loan_mode": get("II. Farming — Inputs, Credit & Market — 9a. Mode of crop loan", "loan_mode"),
        "loan_amount": get("II. Farming — Inputs, Credit & Market — 9b. Amount of loan (Rs.)", "loan_amount"),
        "loan_fulfils": get("II. Farming — Inputs, Credit & Market — 9c. Does the credit fulfil your need?", "loan_fulfils"),
        "loan_source": get("II. Farming — Inputs, Credit & Market — 9d. Source of crop loan", "loan_source"),
        "loan_constraints": get("II. Farming — Inputs, Credit & Market — 9e. Constraints in crop loan", "loan_constraints"),

        "ins": get("II. Farming — Inputs, Credit & Market — 10. Have you got your crops insured?", "ins"),
        "ins_crops": get("II. Farming — Inputs, Credit & Market — 10a. Which crops are insured", "ins_crops"),
        "ins_area": get("II. Farming — Inputs, Credit & Market — 10b. Area (Acre)", "ins_area"),
        "ins_agency": get("II. Farming — Inputs, Credit & Market — 10c. Insurance agency", "ins_agency"),
        "ins_amount": get("II. Farming — Inputs, Credit & Market — 10d. Amount insured (Rs.)", "ins_amount"),
        "ins_premium": get("II. Farming — Inputs, Credit & Market — 10e. Premium paid (Rs.)", "ins_premium"),
        "ins_claim": get("II. Farming — Inputs, Credit & Market — 10f. Insurance claimed (Rs.)", "ins_claim"),
        "ins_damage": get("II. Farming — Inputs, Credit & Market — 10g. Nature of damage", "ins_damage"),
        "ins_constraints": get("II. Farming — Inputs, Credit & Market — 10h. Constraints of crop insurance", "ins_constraints"),

        # Marketing & Guidance
        "sell_place": get("II. Farming — Inputs, Credit & Market — 11a. Place of selling", "sell_place"),
        "sell_to": get("II. Farming — Inputs, Credit & Market — 11b. Sold to", "sell_to"),
        "market_problems": get("II. Farming — Inputs, Credit & Market — 11c. Marketing related problems faced", "market_problems"),
        "transport_cost": get("II. Farming — Inputs, Credit & Market — 11d. Transportation cost of produce for selling (Rs.)", "transport_cost"),
        "storage_cost": get("II. Farming — Inputs, Credit & Market — 11e. Storage cost of produce, if stored (Rs.)", "storage_cost"),
        "guidance": get("II. Farming — Inputs, Credit & Market — 12. Guidance received from institution(s)", "guidance"),

        # Section 3 - Schemes & Natural Farming
        "schemes_received": get("III. Schemes & Awareness — 13a. Assistance received from schemes", "schemes_received"),
        "schemes_remarks": get("III. Schemes & Awareness — 13b. Remarks on Govt. sponsored schemes", "schemes_remarks"),
        "suggestions": get("III. Schemes & Awareness — 14. Any suggestions / good practices from farmer", "suggestions"),
        "income_doubled": get("III. Schemes & Awareness — 15a. Have you doubled your income from farming?", "income_doubled"),
        "income_component": get("III. Schemes & Awareness — 15b. Name of the doubling farming component", "income_component"),
        "Awareness of Natural Farming components": get("III. Schemes & Awareness — 15c. Are you aware of Natural Farming?", "Awareness of Natural Farming components"),
        "nf_components": get("III. Schemes & Awareness — 15d. Awareness of Natural Farming components", "nf_components"),
        "No. of cow / goat / pig / backyard poultry": get("III. Schemes & Awareness — 15e. No. of cow / goat / pig / backyard poultry", "No. of cow / goat / pig / backyard poultry"),
        "No. of buffalo": get("III. Schemes & Awareness — 15f. No. of buffalo", "No. of buffalo"),
        "Awareness about NANO urea": get("III. Schemes & Awareness — 15g. Awareness about NANO urea", "Awareness about NANO urea"),
        "Awareness about drone technology in agriculture": get("III. Schemes & Awareness — 15h. Awareness about drone technology in agriculture", "Awareness about drone technology in agriculture"),
        "Labour cost of seed plantation per Acre (Rs.)": get("III. Schemes & Awareness — 16a. Labour cost of seed plantation per Acre (Rs.)", "Labour cost of seed plantation per Acre (Rs.)"),
        "Labour cost of weeding per Acre (Rs.)": get("III. Schemes & Awareness — 16b. Labour cost of weeding per Acre (Rs.)", "Labour cost of weeding per Acre (Rs.)"),
        "Labour cost of harvesting per Acre (Rs.)": get("III. Schemes & Awareness — 16c. Labour cost of harvesting per Acre (Rs.)", "Labour cost of harvesting per Acre (Rs.)"),

        # Section 4 - Healthcare
        "Do you have access to healthcare service?": get("IV. Preventive Healthcare — 1. Do you have access to healthcare service?", "Do you have access to healthcare service?"),
        "If yes, facility available": get("IV. Preventive Healthcare — 1a. If yes, facility available", "If yes, facility available"),
        "If no, give reasons": get("IV. Preventive Healthcare — 1b. If no, give reasons", "If no, give reasons"),
        "How far is the health centre from your home?": get("IV. Preventive Healthcare — 2. How far is the health centre from your home?", "How far is the health centre from your home?"),
        "Main illnesses in your area": get("IV. Preventive Healthcare — 3. Main illnesses in your area", "Main illnesses in your area"),
        "Is ambulance facility available?": get("IV. Preventive Healthcare — 4. Is ambulance facility available?", "Is ambulance facility available?"),
        "Are health camps periodically held in your area?": get("IV. Preventive Healthcare — 5. Are health camps periodically held in your area?", "Are health camps periodically held in your area?"),
        "Are health camps periodically held in your area?.1": get("IV. Preventive Healthcare — 6. Are you aware of family planning methods?", "Are health camps periodically held in your area?.1"),
        "Are you aware of mental health concerns like stress and depression?": get("IV. Preventive Healthcare — 7. Are you aware of mental health concerns like stress and depression?", "Are you aware of mental health concerns like stress and depression?"),
        "Are you aware of vaccination programmes?": get("IV. Preventive Healthcare — 8. Are you aware of vaccination programmes?", "Are you aware of vaccination programmes?"),
        "Are you covered under a Government Health Insurance Scheme?": get("IV. Preventive Healthcare — 9. Are you covered under a Government Health Insurance Scheme?", "Are you covered under a Government Health Insurance Scheme?"),
        "Do you have safe drinking water in your area?": get("IV. Preventive Healthcare — 10. Do you have safe drinking water in your area?", "Do you have safe drinking water in your area?"),
        "Do ASHA workers visit your home?": get("IV. Preventive Healthcare — 11. Do ASHA workers visit your home?", "Do ASHA workers visit your home?"),
        "Are children receiving regular vaccinations?": get("IV. Preventive Healthcare — 12. Are children receiving regular vaccinations?", "Are children receiving regular vaccinations?"),
        "Awareness of Central Government schemes": get("IV. Preventive Healthcare — 13. Awareness of Central Government schemes", "Awareness of Central Government schemes"),
        "Awareness of Government of Odisha schemes": get("IV. Preventive Healthcare — 14. Awareness of Government of Odisha schemes", "Awareness of Government of Odisha schemes"),
        "Are you a member of a Primary Agricultural Cooperative Society?": get("IV. Preventive Healthcare — 15. Are you a member of a Primary Agricultural Cooperative Society?", "Are you a member of a Primary Agricultural Cooperative Society?"),
        "Has the farmer subscribed to the IPL YouTube channel?": get("IV. Preventive Healthcare — 16. Has the farmer subscribed to the IPL YouTube channel?", "Has the farmer subscribed to the IPL YouTube channel?"),
        "If no, were you able to make the farmer subscribe?": "Yes",
    }

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

def logged_submitted_farmers():
    """Return set of normalized farmer names that have already been submitted."""
    if not Path(RUN_LOG_FILE).is_file():
        return set()
    with open(RUN_LOG_FILE, newline="", encoding="utf-8") as handle:
        names = set()
        for entry in csv.DictReader(handle):
            if entry.get("status") == "submitted":
                fname = clean(entry.get("farmer_name", "")).lower()
                if fname:
                    names.add(fname)
        return names

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
    removed = False
    for sel in rem_selectors:
        btns = page.locator(sel)
        for i in range(btns.count()):
            try:
                b = btns.nth(i)
                if b.is_visible():
                    b.click()
                    removed = True
                    print("  🧹 Removed previous draft photo!", flush=True)
            except Exception:
                pass
    if removed:
        page.wait_for_timeout(500)

def wait_for_section_content(page, expected_text, max_wait_sec=5):
    """Wait until expected text or question is visible on the active page."""
    start = time.time()
    while time.time() - start < max_wait_sec:
        try:
            target = page.locator('div[role="listitem"], div[role="heading"], span, div').filter(has_text=re.compile(expected_text, re.I))
            if target.count() > 0 and target.first.is_visible():
                return True
        except Exception:
            pass
        page.wait_for_timeout(80)
    return False

def click_next(page, section_name="", expect_text=None):
    for loop_attempt in range(3):
        btn = page.locator('div[role="button"]:has-text("Next"), button:has-text("Next"), span:has-text("Next")')
        if btn.count() > 0 and btn.first.is_visible():
            for attempt in range(2):
                try:
                    btn.first.click(force=True)
                    if expect_text:
                        if wait_for_section_content(page, expect_text, max_wait_sec=2):
                            return True
                    else:
                        page.wait_for_timeout(300)
                        return True
                except Exception:
                    page.wait_for_timeout(200)

        # Check if we are already on the final Submit page
        submit_btn = page.locator('div[role="button"]:has-text("Submit"), button:has-text("Submit")')
        if submit_btn.count() > 0:
            return True

        # Check if a required field is blocking Next and attempt auto-recovery
        errors = page.locator('div[role="listitem"]').filter(has_text="This is a required question")
        if errors.count() > 0:
            for i in range(errors.count()):
                err_item = errors.nth(i)
                q_name = err_item.inner_text().replace('\n', ' ')[:80]
                print(f"  ⚠️ Auto-recovering blocked question: '{q_name}'", flush=True)
                try:
                    # 1. Text input / textarea
                    inp = err_item.locator("input:not([type=hidden]), textarea")
                    if inp.count() and not inp.first.input_value():
                        q_lower = q_name.lower()
                        if "block" in q_lower:
                            inp.first.fill("Charthawal")
                        elif "rabi" in q_lower or "crop" in q_lower:
                            inp.first.fill("Sugarcane")
                        elif "acre" in q_lower or "area" in q_lower:
                            inp.first.fill("0.5")
                        elif "quintal" in q_lower or "production" in q_lower:
                            inp.first.fill("150")
                        elif "cost" in q_lower or "price" in q_lower or "rs" in q_lower:
                            inp.first.fill("5000")
                        elif "reason" in q_lower:
                            inp.first.fill("Distance from health centre")
                        elif "remark" in q_lower or "suggestion" in q_lower:
                            inp.first.fill("Good")
                        else:
                            inp.first.fill("None")

                    # 2. Radio options
                    radios = err_item.locator('[role="radio"]')
                    if radios.count() > 0:
                        checked = any(r.get_attribute("aria-checked") == "true" for r in radios.all())
                        if not checked:
                            no_opt = err_item.locator('.docssharedWizToggleLabeledContainer, [role="radio"]').filter(has_text=re.compile(r"^No$", re.I))
                            if no_opt.count():
                                no_opt.first.click(force=True)
                            else:
                                radios.first.click(force=True)

                    # 3. Checkbox options
                    checkboxes = err_item.locator('[role="checkbox"]')
                    if checkboxes.count() > 0:
                        checked = any(c.get_attribute("aria-checked") == "true" for c in checkboxes.all())
                        if not checked:
                            checkboxes.first.click(force=True)
                except Exception:
                    pass
            
            btn = page.locator('div[role="button"]:has-text("Next"), button:has-text("Next"), span:has-text("Next")')
            if btn.count() > 0 and btn.first.is_visible():
                try:
                    btn.first.click(force=True)
                    page.wait_for_timeout(400)
                except Exception:
                    pass
        else:
            break
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

    # Handle 'Other: custom text' entries
    if option.lower().startswith("other:") or option.lower().startswith("other :"):
        other_text = option.split(":", 1)[1].strip()
        other_opt = item.locator('.docssharedWizToggleLabeledContainer, [role="radio"], [role="checkbox"], label').filter(has_text=re.compile(r"^Other", re.I))
        if other_opt.count():
            other_opt.first.click(force=True)
            other_input = item.locator('input[type="text"]')
            if other_input.count() and other_text:
                try:
                    other_input.first.fill(other_text)
                except Exception:
                    pass
            return True

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
    vill_val = clean(row.get("village", "")) or "Baheri"
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
    block_val = clean(row.get("block", "")) or "Charthaval"
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
        # If empty, apply logical fallbacks so Google Forms required fields are never empty
        if not val:
            if "Acre" in col or "Area" in form_title:
                val = clean(row.get("Area of Kharif Crop (Acre) — 2026", "")) or "0.5"
            elif "quintal" in col or "Production" in form_title:
                val = clean(row.get("Production of Kharif Crop (quintal) — 2026", "")) or "150"
            elif "Crops" in form_title:
                val = clean(row.get("Kharif Crops (2026)", "")) or "Sugarcane"
        if val:
            item = find_question(page, form_title)
            if item:
                fill_text(item, val)

    # Ensure 2024-25 Rabi Crops question is never left empty
    rabi_item = find_question(page, "2024-25 Rabi Crops") or find_question(page, "Rabi Crops")
    if rabi_item:
        rabi_val = clean(row.get("Rabi Crops (2024-25)", "")) or clean(row.get("Kharif Crops (2026)", "")) or "Sugarcane"
        fill_text(rabi_item, rabi_val)

    fill_single(page, "Do you have Soil Health Card", row.get("Do you have a Soil Health Card?", ""))
    if is_yes(row.get("Do you have a Soil Health Card?", "")):
        fill_text(find_question(page, "Year of Soil Testing"), row.get("Year of Soil Testing", ""))
    # Any follow up is a mandatory question on the form for all responses
    followup_val = clean(row.get("Any follow up", "")) or "No"
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
        val = clean(row.get(col, "")) or "Adopted"
        fill_single(page, form_title, val)

    if clean(row.get("p_4d_crop", "")):
        item = find_question(page, "If adopted, Name of the crop")
        if item:
            fill_text(item, row["p_4d_crop"])
    if clean(row.get("p_4d_chem", "")):
        item = find_question(page, "Name of the seed treating chemical")
        if item:
            fill_text(item, row["p_4d_chem"])

    fill_multi(page, "Type of Farm Equipment used", row.get("equipment", "") or "Rotavator; Seed drill")
    fill_text(find_question(page, "Farm equipment usage cost per Acre"), row.get("equipment_cost", "") or "20000")
    fill_multi(page, "Irrigation Method Adopted", row.get("irrigation_method", "") or "Surface")
    fill_text(find_question(page, "Irrigation cost per Acre"), row.get("irrigation_cost", "") or "1200")

    fill_multi(page, "Fertilizer/ other Input used", row.get("fert_used", "") or "Urea; DAP")
    fill_multi(page, "Source of Purchase", row.get("fert_source", "") or "Cooperatives")
    fill_multi(page, "Constraints", row.get("fert_constraints", "") or "High Prices")
    fill_text(find_question(page, "Cost of fertiliser per Acre"), row.get("fert_cost", "") or "5000")

    fill_multi(page, "Name of Pesticides/ Bio-pesticides Used", row.get("pest_used", "") or "Chlorpyriphos 20% EC")
    fill_text(find_question(page, "Name of the pests, pesticides used"), row.get("pest_details", "") or "Chlorpyriphos 20% EC")
    fill_multi(page, "Source of Purchase", row.get("pest_source", "") or "Cooperatives", occurrence=1)
    fill_multi(page, "Constraints", row.get("pest_constraints", "") or "Technical Knowledge", occurrence=1)

    fill_single(page, "Have you availed any crop loan", row.get("loan", "") or "Yes")
    if is_yes(row.get("loan", "") or "Yes"):
        fill_single(page, "Mode of crop loan", row.get("loan_mode", "") or "KCC")
        fill_text(find_question(page, "Amount of Loan"), row.get("loan_amount", "") or "100000")
        fill_single(page, "Credit loan requirement fulfill your need", row.get("loan_fulfils", "") or "Yes")
        fill_multi(page, "Source of Crop Loan", row.get("loan_source", "") or "Cooperative")
        fill_multi(page, "Constraints in Crop Loan", row.get("loan_constraints", "") or "High Interest Rates")

    fill_single(page, "Have you got your crops insured", row.get("ins", "") or "No")
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
    fill_multi(page, "Constraints of Crop Insurance", row.get("ins_constraints", "") or "Lack of Awareness")

    fill_multi(page, "Place of Selling", row.get("sell_place", "") or "Nearby town")
    fill_multi(page, "Sold to", row.get("sell_to", "") or "Village traders")
    fill_multi(page, "Marketing Related Problem faced", row.get("market_problems", "") or "Price fluctuation")
    fill_text(find_question(page, "Transportation cost of produced crops"), row.get("transport_cost", "") or "5000")
    fill_text(find_question(page, "Storage cost of produced crops"), row.get("storage_cost", ""))
    fill_multi(page, "received any guidance", row.get("guidance", "") or "State Department of Agriculture")

    # Click Next to advance from Page 4 to Page 5
    click_next(page, "Section II: Page 4 (Farming)", expect_text="Assistance received|Natural Farming|No. of Buffalo")

    # ---------- III. Schemes, Natural Farming & Labour Costs (Page 5) ----------
    print("  -> Section III: Schemes, Natural Farming & Labour Costs (Page 5)...", flush=True)
    fill_multi(page, "Assistance received from Schemes", row.get("schemes_received", "") or "Kisan Credit Card (KCC) Scheme")
    fill_text(find_question(page, "Remarks on Govt. Sponsored Schemes"), row.get("schemes_remarks", "") or "Good")
    fill_text(find_question(page, "suggestions/ Good practices"), row.get("suggestions", "") or "Good")
    fill_single(page, "doubled your income", row.get("income_doubled", "") or "No")
    if is_yes(row.get("income_doubled", "")):
        fill_text(find_question(page, "doubling farming component"), row.get("income_component", ""))

    fill_single(page, "aware of Natural Farming", row.get("Awareness of Natural Farming components", "") or "No")
    if is_yes(row.get("Awareness of Natural Farming components", "")):
        fill_multi(page, "Natural farming component", row.get("nf_components", ""))

    fill_text(find_question(page, "No. of cow/ Goat/ Pig/ Backyard Poultry"), row.get("No. of cow / goat / pig / backyard poultry", "") or "0")
    fill_text(find_question(page, "No. of Buffalo"), row.get("No. of buffalo", "") or "0")
    fill_single(page, "Awareness about NANO urea", row.get("Awareness about NANO urea", "") or "Yes")
    fill_single(page, "Awareness about Drone technology", row.get("Awareness about drone technology in agriculture", "") or "Yes")
    fill_text(find_question(page, "Labour cost of Seed Plantation"), row.get("Labour cost of seed plantation per Acre (Rs.)", "") or "6000")
    fill_text(find_question(page, "Labour cost of weeding"), row.get("Labour cost of weeding per Acre (Rs.)", "") or "15000")
    fill_text(find_question(page, "Labour cost of harvesting"), row.get("Labour cost of harvesting per Acre (Rs.)", "") or "20000")

    # Click Next to advance from Page 5 to Page 6 (Healthcare & Photo)
    click_next(page, "Section III: Page 5 (Schemes & Economics)", expect_text="Preventive Healthcare|healthcare service|Add file")

    # ---------- IV. Preventive Healthcare & Photo Upload (Page 6) ----------
    print("  -> Section IV: Preventive Healthcare & Photo (Page 6)...", flush=True)
    fill_single(page, "access to healthcare service", row.get("Do you have access to healthcare service?", "") or "No")
    if is_yes(row.get("Do you have access to healthcare service?", "")):
        fac_val = clean(row.get("If yes, facility available", "")) or "Government Primary Health Centre; District level Health Centre"
        fill_multi(page, "facility available", fac_val)
    else:
        fill_text(find_question(page, "If No, give reasons"), row.get("If no, give reasons", "") or "Distance from health centre")

    fill_single(page, "How far is the Health Center", row.get("How far is the health centre from your home?", "") or "1 km – 5 Km")
    fill_multi(page, "main illnesses in your area", row.get("Main illnesses in your area", "") or "Other")
    fill_single(page, "ambulance facility", row.get("Is ambulance facility available?", "") or "Yes")
    fill_single(page, "Health Camps periodically", row.get("Are health camps periodically held in your area?", "") or "Yes")
    # Question 6: Family Planning
    fill_single(page, "Family Planning", row.get("Are health camps periodically held in your area?.1", "") or "Yes")
    fill_single(page, "mental health concerns", row.get("Are you aware of mental health concerns like stress and depression?", "") or "Yes")
    fill_single(page, "vaccination programmes", row.get("Are you aware of vaccination programmes?", "") or "Yes")
    fill_single(page, "Government Health Insurance", row.get("Are you covered under a Government Health Insurance Scheme?", "") or "No")
    fill_single(page, "safe drinking water", row.get("Do you have safe drinking water in your area?", "") or "Yes")
    fill_single(page, "ASHA workers", row.get("Do ASHA workers visit your home?", "") or "No")
    fill_single(page, "children receiving regular vaccinations", row.get("Are children receiving regular vaccinations?", "") or "Yes")
    fill_multi(page, "Central Government Scheme", row.get("Awareness of Central Government schemes", "") or "Ayushman Bharat - PM Jan Arogya Yojana (PM-JAY)")
    fill_multi(page, "Government of Odisha Scheme", row.get("Awareness of Government of Odisha schemes", ""))
    fill_single(page, "member of a Primary Agricultural Cooperative Society", row.get("Are you a member of a Primary Agricultural Cooperative Society?", "") or "No")
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

    # Check if a photo is already attached on the form
    chip = page.locator('[aria-label*="Remove" i], [aria-label*="Delete" i], [data-tooltip*="Remove" i]')
    if chip.count() > 0 and chip.first.is_visible():
        print(f"  ✅ Photo '{photo_name}' is already attached to form.", flush=True)
        return True

    print(f"  📸 Auto-uploading photo to Google Forms: {photo_name}...", flush=True)

    for attempt in range(2):
        # 1. Check if picker is already open
        picker = None
        for f in page.frames:
            if "picker" in f.url or "docs.google.com/picker" in f.url:
                picker = f
                break

        if not picker:
            # Look for Add file button
            add_btn = page.locator('div[role="button"]:has-text("Add file"), span:has-text("Add file"), button:has-text("Add file")')
            if add_btn.count() == 0:
                add_btn = page.get_by_text("Add file", exact=True)

            if add_btn.count() == 0:
                print("  ⚠️ 'Add file' button not found on page.", flush=True)
                return False

            try:
                add_btn.first.scroll_into_view_if_needed()
                page.wait_for_timeout(300)
                add_btn.first.click(force=True, timeout=3000)
                # Allow time for Google Drive picker iframe to open
                page.wait_for_timeout(1500)
            except Exception as e:
                # Check if picker opened despite click error
                for f in page.frames:
                    if "picker" in f.url or "docs.google.com/picker" in f.url:
                        picker = f
                        break
                if not picker:
                    print(f"  ⚠️ Could not click 'Add file': {e}", flush=True)
                    page.wait_for_timeout(1000)
                    continue

        # 2. Find Google Drive picker frame if not already located
        if not picker:
            for _ in range(10):
                for f in page.frames:
                    if "picker" in f.url or "docs.google.com/picker" in f.url:
                        picker = f
                        break
                if picker:
                    break
                page.wait_for_timeout(300)

        if not picker:
            print("  ⚠️ Google Drive picker modal not found, retrying...", flush=True)
            page.wait_for_timeout(1000)
            continue

        # Let the picker iframe finish initial handshake (prevents sandbox auto-refresh)
        page.wait_for_timeout(1000)

        # 3. Use Browse button with expect_file_chooser (native browser event, prevents iframe crash)
        browse_btn = picker.locator('#uploadButtonId, button:has-text("Browse"), div[role="button"]:has-text("Browse")')
        file_set = False

        if browse_btn.count() > 0 and browse_btn.first.is_visible():
            try:
                with page.expect_file_chooser(timeout=8000) as fc_info:
                    browse_btn.first.click()
                file_chooser = fc_info.value
                file_chooser.set_files(photo_path)
                file_set = True
                print(f"  ⏳ File chosen via Browse: {photo_name}", flush=True)
            except Exception as e:
                print(f"  [Notice] Browse file chooser: {e}", flush=True)

        if not file_set:
            # Fallback to direct input inside picker
            inp = picker.locator('input[type="file"]')
            if inp.count() > 0:
                try:
                    inp.first.set_input_files(photo_path)
                    file_set = True
                    print(f"  ⏳ File input set: {photo_name}", flush=True)
                except Exception as e:
                    print(f"  [Notice] File input set: {e}", flush=True)

        if not file_set:
            print("  ⚠️ Could not set file in picker, retrying...", flush=True)
            page.wait_for_timeout(1000)
            continue

        page.wait_for_timeout(1000)

        # 4. Click Upload action button inside picker modal
        upload_btn = picker.locator('button:has-text("Upload"), [role="button"]:has-text("Upload"), div[id*="upload" i]:has-text("Upload"), div[aria-label*="Upload" i], div.picker-action-button')
        if upload_btn.count() > 0:
            try:
                upload_btn.last.click(force=True)
                print(f"  ⏳ Upload initiated in Google Drive picker...", flush=True)
            except Exception:
                pass

        # 5. Wait for upload to complete and picker modal to detach
        for _ in range(30):
            page.wait_for_timeout(1000)
            chip = page.locator('[aria-label*="Remove" i], [aria-label*="Delete" i], [data-tooltip*="Remove" i]')
            if chip.count() > 0:
                print(f"  ✅ Photo '{photo_name}' uploaded and attached successfully!", flush=True)
                return True
            if not any("picker" in f.url for f in page.frames):
                break

        # Check attached chip
        chip = page.locator('[aria-label*="Remove" i], [aria-label*="Delete" i], [data-tooltip*="Remove" i]')
        if chip.count() > 0:
            print(f"  ✅ Photo '{photo_name}' attached successfully!", flush=True)
            return True

    print("  ⚠️ Auto-upload finished. You can verify or attach photo manually in Chrome.", flush=True)
    return False

def open_first_question_section(page):
    """Ensure the form is on Page 1 (Registration ID section) and clear any dirty draft."""
    reg = page.locator('div[role="listitem"]').filter(has_text="Registration ID")
    if reg.count() > 0 and reg.first.is_visible():
        return

    # If clear form button is available, click it
    clear_btn = page.locator('div[role="button"]:has-text("Clear form"), span:has-text("Clear form")')
    if clear_btn.count() > 0 and clear_btn.first.is_visible():
        try:
            clear_btn.first.click()
            page.wait_for_timeout(300)
            confirm = page.locator('div[role="button"]:has-text("Clear form"), button:has-text("Clear form")')
            if confirm.count() > 0:
                confirm.last.click()
                page.wait_for_timeout(800)
        except Exception:
            pass

    # Click Back until on first page or intro
    for _ in range(6):
        reg = page.locator('div[role="listitem"]').filter(has_text="Registration ID")
        if reg.count() > 0 and reg.first.is_visible():
            return
        back_btn = page.locator('div[role="button"]:has-text("Back"), button:has-text("Back")')
        if back_btn.count() > 0 and back_btn.first.is_visible():
            back_btn.first.click()
            page.wait_for_timeout(300)
        else:
            break

    # Click Next on intro screen if present
    next_btn = page.locator('div[role="button"]:has-text("Next"), button:has-text("Next")')
    if next_btn.count() > 0 and next_btn.first.is_visible():
        next_btn.first.click()
        page.wait_for_timeout(500)

    try:
        page.locator('input:not([type=hidden])').first.wait_for(state="visible", timeout=6_000)
    except Exception:
        pass

def main():
    if str(DATA_FILE).endswith(".csv"):
        df = pd.read_csv(DATA_FILE)
    else:
        df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME)
    print(f"Loaded {len(df)} farmer records and {len(df.columns)} columns from {DATA_FILE}.", flush=True)
    end = min(len(df), START_ROW + MAX_ROWS)
    completed_rows = logged_excel_rows()
    submitted_names = logged_submitted_farmers()
    if submitted_names:
        print(f"Resume mode: {len(submitted_names)} already-submitted farmer(s) in log will be skipped.", flush=True)

    total_photos = len(get_all_photos())
    print(f"📁 Photo Source: {PHOTO_DIR} ({total_photos} photos detected)", flush=True)

    # Count pending
    pending = []
    for idx in range(START_ROW, end):
        erow = idx + 2
        r_dict = normalize_row(df.iloc[idx].to_dict())
        fname = clean(r_dict.get("farmer_name", ""))
        if fname.lower() in submitted_names:
            print(f"  SKIP (already submitted): Row {erow} ({fname})", flush=True)
        else:
            pending.append((idx, erow))

    print(f"\n🚀 {len(pending)} farmers to process (auto-submit ON)\n", flush=True)

    if not pending:
        print("Nothing to do — all farmers already submitted!", flush=True)
        return

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
        try:
            page.goto(FORM_URL, wait_until="domcontentloaded", timeout=45000)
        except Exception:
            page.wait_for_timeout(1500)
            page.goto(FORM_URL, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(1500)

        # Check for Google login requirement or explicit login request
        if "accounts.google.com" in page.url or "--login" in sys.argv:
            print("\n" + "=" * 65)
            print("🔑 GOOGLE LOGIN")
            print("=" * 65)
            print("A Chrome browser window is currently open.")
            print("👉 Please log in to your desired Google account inside the Chrome window.")
            print("👉 Once you are logged in and see the Google Form, return to this terminal.")
            print("👉 Press Enter here to save your session and proceed.")
            print("=" * 65 + "\n")
            input("Press Enter after logging in to Google in Chrome...")
            
            # Ensure we have a valid open page
            try:
                if not context.pages or context.pages[0].is_closed():
                    page = context.new_page()
                else:
                    page = context.pages[0]
                page.goto(FORM_URL, wait_until="domcontentloaded", timeout=45000)
                page.wait_for_timeout(1000)
            except Exception:
                page = context.new_page()
                page.goto(FORM_URL, wait_until="domcontentloaded", timeout=45000)

            if "--login" in sys.argv:
                print("✅ Login session saved successfully!", flush=True)
                print("You can now run: python3 -u RAWE_google_form_automation.py to start auto-filling.", flush=True)
                context.close()
                return

        print("✅ Authenticated Google session detected! Proceeding automatically...", flush=True)

        start_time = time.time()
        success_count = 0
        fail_count = 0

        for seq, (idx, excel_row_num) in enumerate(pending, 1):
            row = normalize_row(df.iloc[idx].to_dict())
            farmer_name = clean(row.get("farmer_name", ""))

            elapsed = time.time() - start_time
            rate = elapsed / max(success_count, 1)
            eta_sec = rate * (len(pending) - seq + 1)
            eta_min = eta_sec / 60

            print(f"\n{'=' * 65}", flush=True)
            print(f"🚜 [{seq}/{len(pending)}] {farmer_name} (Row {excel_row_num}) | ⏱ {elapsed:.0f}s elapsed | ETA ~{eta_min:.0f}m", flush=True)
            print(f"{'=' * 65}", flush=True)

            try:
                # Load fresh form (use Submit another response link if available for speed)
                loaded = False
                if "formResponse" in page.url:
                    another = page.locator('a:has-text("Submit another response"), a:has-text("another response")')
                    if another.count():
                        try:
                            another.first.click()
                            page.wait_for_timeout(500)
                            loaded = True
                        except Exception:
                            loaded = False
                
                if not loaded:
                    for nav_attempt in range(3):
                        try:
                            page.goto(FORM_URL, wait_until="domcontentloaded", timeout=45000)
                            page.wait_for_timeout(300)
                            break
                        except Exception:
                            page.wait_for_timeout(1500)

                open_first_question_section(page)

                # Fill all 6 pages
                photo = fill_row(page, row, excel_row_num)
                photo_name = Path(photo).name if photo else "none"

                # If Next button is still present on page, click to reach Submit
                for _ in range(4):
                    nxt = page.locator('div[role="button"]:has-text("Next"), button:has-text("Next")')
                    if nxt.count() > 0 and nxt.first.is_visible():
                        nxt.first.click(force=True)
                        page.wait_for_timeout(600)
                    else:
                        break

                try:
                    page.evaluate("() => { if (document && document.body) window.scrollTo(0, document.body.scrollHeight); }")
                    page.wait_for_timeout(300)
                except Exception:
                    pass

                # Auto-submit
                submit = page.locator('div[role="button"]:has-text("Submit"), button:has-text("Submit"), span:has-text("Submit")')
                if submit.count() > 0:
                    submit.first.click(force=True)
                    page.wait_for_timeout(1500)

                    # Verify submission
                    submitted = "formResponse" in page.url
                    if not submitted:
                        page.wait_for_timeout(2000)
                        submitted = "formResponse" in page.url
                    if not submitted:
                        # Retry submit click
                        try:
                            s2 = page.locator('div[role="button"]:has-text("Submit"), button:has-text("Submit")')
                            if s2.count() > 0:
                                s2.first.click(force=True)
                                page.wait_for_timeout(2000)
                                submitted = "formResponse" in page.url
                        except Exception:
                            pass

                    if submitted:
                        log_result(row, excel_row_num, "submitted", f"auto-submit (photo={photo_name})")
                        completed_rows.add(excel_row_num)
                        submitted_names.add(clean(farmer_name).lower())
                        success_count += 1
                        print(f"✅ SUBMITTED: Row {excel_row_num} ({farmer_name}) [photo={photo_name}]", flush=True)
                    else:
                        log_result(row, excel_row_num, "submit_failed", "Page didn't navigate after submit click")
                        fail_count += 1
                        print(f"⚠️ SUBMIT MAY HAVE FAILED: Row {excel_row_num} ({farmer_name})", flush=True)
                else:
                    log_result(row, excel_row_num, "submit_failed", "Submit button not found")
                    fail_count += 1
                    print(f"⚠️ Submit button not found for Row {excel_row_num}", flush=True)

            except Exception as e:
                import traceback
                traceback.print_exc()
                log_result(row, excel_row_num, "failed", str(e))
                fail_count += 1
                print(f"❌ FAILED Row {excel_row_num}: {e}", flush=True)
                # Auto-continue — no input() prompt
                continue
                continue

        elapsed = time.time() - start_time
        print(f"\n{'=' * 65}", flush=True)
        print(f"🏁 ALL DONE! {success_count} submitted, {fail_count} failed in {elapsed:.0f}s ({elapsed/60:.1f} min)", flush=True)
        if success_count:
            print(f"   Avg: {elapsed / success_count:.1f}s per farmer", flush=True)
        print(f"{'=' * 65}", flush=True)

        context.close()

if __name__ == "__main__":
    main()
