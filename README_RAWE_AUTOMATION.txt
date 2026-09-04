RAWE Google Form Automation — FIRST TEST

1. Install Python 3.10+ on a Windows/Mac/Linux computer.
2. Put these files in one folder:
   - RAWE_google_form_automation.py
   - requirements.txt
   - Anand_IPL_RAWE_2026.xlsx
3. Open Terminal/Command Prompt in that folder.
4. Run:
      pip install -r requirements.txt
      playwright install chromium
5. Run:
      python RAWE_google_form_automation.py

IMPORTANT:
- The script is currently in DRY_RUN mode.
- It fills only the first farmer and DOES NOT click Submit.
- A browser window will open. Sign in to the Google account that has permission to submit the form, if asked.
- Review the filled form before changing DRY_RUN=False.
- The geotag photo upload is not configured yet. We need the photo files and their relationship to each Excel row before batch submission.
- Do not run all 100 records until the first dry run is visually verified.

After the first test works, we will configure:
- photo folder / photo-to-farmer mapping
- exact checkbox/radio mappings for any remaining mismatches
- batch size and resume logging

LOCAL BATCH MODE (after the dry run has been reviewed)

1. Create a photo folder and set PHOTO_DIR in RAWE_google_form_automation.py.
   Name each photo using its registration ID, for example:
      ICRO_2026_20260800788.jpg
   Or add explicit PHOTO_MAP entries for files with other names.
2. Set DRY_RUN = False and MAX_ROWS = 100.
3. Run the script again. It will refuse to submit if any farmer in the run
   has no matched photo.
4. Progress is written to rawe_submission_log.csv. Re-running the script skips
   registration IDs recorded with status "submitted", so it can safely resume.

The script never submits in DRY_RUN mode. Do not delete the log file while a
batch is in progress.
