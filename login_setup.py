"""
One-Time Google Sign-In Helper for RAWE Automation
===================================================
Launches your real Google Chrome browser with 'google_form_browser_profile'
WITHOUT automation flags so Google will NOT block your login with
"This browser or app may not be secure".
"""

import subprocess
import sys
from pathlib import Path

FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeAjfujYObtb_K79YzlOAOAMeo8dclD2C7jjDpyl3UMZnQ8dg/viewform"
PROFILE_DIR = Path(__file__).resolve().parent / "google_form_browser_profile"
CHROME_PATH = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

def main():
    print("=" * 65)
    print("Launching real Google Chrome without automation flags...")
    print("=" * 65)

    if not Path(CHROME_PATH).exists():
        print(f"Error: Google Chrome was not found at {CHROME_PATH}")
        sys.exit(1)

    cmd = [
        CHROME_PATH,
        f"--user-data-dir={PROFILE_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        FORM_URL,
    ]

    print("\nA dedicated Google Chrome window is opening now.")
    print("1. Sign in with your Google account in that Chrome window.")
    print("2. Once you can see the RAWE Google Form questions, close that Chrome window.")
    print("3. Return here and press ENTER.")
    print("-" * 65)

    proc = subprocess.Popen(cmd)
    input("\nPress ENTER after you have signed in and closed the Chrome window...")

    try:
        proc.terminate()
    except Exception:
        pass

    print("\n✅ Session saved in 'google_form_browser_profile'!")
    print("You can now run: python3 RAWE_google_form_automation.py")

if __name__ == "__main__":
    main()

