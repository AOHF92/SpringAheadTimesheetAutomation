"""
Step 1 – Fetch worked days from SpringAhead using Playwright.

Setup (run once per environment):
    pip install playwright python-dotenv
    playwright install

Usage:
    python springahead_step1_fetch.py

Behavior:
    - If .env with SPRINGAHEAD_* vars exists and is complete, use it.
    - Otherwise, prompt for credentials.
    - Optionally saves new .env for next runs.
    - Logs into SpringAhead, clicks "Add Time",
      scrapes current week's days with hours > 0,
      and prints them + saves to JSON.
"""

import os
from pathlib import Path
from getpass import getpass
import json
import sys

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from dotenv import load_dotenv

def get_app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

LOGIN_URL = (
    "https://my.springahead.com/go/Account/Logon"
    "?ReturnUrl=%2Fvt%2Fgo%3FHome%26tokenid%3Dvte"
)

APP_ROOT = get_app_root()
ENV_PATH = APP_ROOT / "MyCreds.env"
OUTPUT_JSON = APP_ROOT / "springahead_current_week.json"


def load_credentials():
    """
    Load SpringAhead credentials in this order:
      1. Environment variables
      2. .env file
      3. Interactive prompt (only if still incomplete)
    """
    """
    # Debug tracking
    print(f"[DEBUG] CWD: {Path.cwd()}")
    print(f"[DEBUG] Looking for creds at: {ENV_PATH} (exists={ENV_PATH.exists()})")
    """
    # --- 1) Environment variables (GUI sets these) ---
    company = os.getenv("SPRINGAHEAD_COMPANY") or ""
    username = os.getenv("SPRINGAHEAD_USERNAME") or ""
    password = os.getenv("SPRINGAHEAD_PASSWORD") or ""

    if company and username and password:
        return {
            "company": company,
            "username": username,
            "password": password,
        }

    # --- 2) Try .env file ---
    if ENV_PATH.exists():
        load_dotenv(dotenv_path=ENV_PATH)

        company = os.getenv("SPRINGAHEAD_COMPANY") or company
        username = os.getenv("SPRINGAHEAD_USERNAME") or username
        password = os.getenv("SPRINGAHEAD_PASSWORD") or password

        if company and username and password:
            return {
                "company": company,
                "username": username,
                "password": password,
            }

    if getattr(sys, "frozen", False):
        # Running inside a bundled EXE: no interactive console available
        raise RuntimeError(
        "No credentials found in MyCreds.env or environment variables. "
        "Please create MyCreds.env next to the EXE or use the GUI fields."
       )
    print("No complete .env or environment variables found. Please enter your SpringAhead credentials.")
    # If you're running under Gooey, this will hang, so for GUI usage
    # you should ALWAYS fill the fields or have a valid .env.
    company = company or input("Company: ").strip()
    username = username or input("Login name: ").strip()
    password = password or getpass("Password: ")

    return {
        "company": company,
        "username": username,
        "password": password,
    }

def fetch_worked_days(creds, headless=True):
    results = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page()

        print("Opening login page...")
        page.goto(LOGIN_URL, wait_until="domcontentloaded")

        # --- LOGIN ---
        print("Filling login form...")

        # Scope to the main login form only
        page.wait_for_selector("#login_body input#CompanyLogin", timeout=15000)
        page.wait_for_selector("#login_body input#UserName", timeout=15000)
        page.wait_for_selector("#login_body input#Password", timeout=15000)

        page.locator("#login_body input#CompanyLogin").fill(creds["company"])
        page.locator("#login_body input#UserName").fill(creds["username"])
        page.locator("#login_body input#Password").fill(creds["password"])

        page.get_by_role("button", name="Log In").click()
        
        # give the page a moment to redraw after login
        page.wait_for_load_state("networkidle")

        # Look for login-error banner
        # Use the visible text from the page; no extra quotes needed
        error_banner = page.locator("text=Login information entered is invalid. Please try again.")

        if error_banner.is_visible():
            # Optional: screenshot for debugging
            page.screenshot(path="springahead_login_error.png", full_page=True)

            raise RuntimeError(
                "SpringAhead login failed: login information is invalid. "
                "Please check your company, username, or password (MyCreds.env / GUI)."
            )

        # --- HOME PAGE (Add Time) ---
        try:
            page.get_by_text("Add Time", exact=True).wait_for(timeout=15000)
        except PlaywrightTimeoutError:
            browser.close()
            raise RuntimeError(
                "Could not find 'Add Time' after logging in. "
                "Check credentials or if the UI changed."
            )


        print("Clicking 'Add Time' to open current timecard...")
        page.get_by_text("Add Time", exact=True).click()

        # --- TIME ENTRY PAGE ---
        try:
            page.get_by_text("Enter Time for", exact=False).wait_for(timeout=15000)
        except PlaywrightTimeoutError:
            browser.close()
            raise RuntimeError(
                "Time entry page did not load (no 'Enter Time for' found)."
            )

        # Optional: visual delay so the UI actually fully loads
        page.wait_for_timeout(3000)
        # --- Switch to List view (Week view loads by default with no cookies) ---
        print("Switching to List view...")
        page.get_by_text("List", exact=True).click()
        page.wait_for_timeout(3000)

        print("Waiting for timecard table to load...")
        page.wait_for_selector("table.timedayTable", timeout=20000)

        print("Scraping worked days from the timecard...")

        rows = page.locator("table.timedayTable tr.timeRow")
        row_count = rows.count()
        print(f"Found {row_count} time row(s) on the page.")

        for i in range(row_count):
            row = rows.nth(i)

            date_text = row.locator(".timedayDate").inner_text().strip()
            project_text = row.locator("span.timedayProject").inner_text().strip()
            type_text = row.locator("td.timedayType .timedayType").inner_text().strip()
            hours_text = row.locator("td.timedayHours").inner_text().strip()

            if not hours_text:
                continue

            try:
                hours_val = float(hours_text)
            except ValueError:
                print(f"Skipping row with non-numeric hours: {hours_text!r}")
                continue

            if hours_val <= 0:
                continue

            entry = {
                "date": date_text,
                "hours": hours_val,
                "project": project_text,
                "type": type_text,
            }
            results.append(entry)

        browser.close()

    return results


def main():
    creds = load_credentials()

    # Decide headless mode from env (default: headless ON)
    #
    # SPRINGAHEAD_HEADLESS values treated as "off" (headed):
    #   0, "false", "no", "off"  (case-insensitive)
    #
    # Anything else (or unset) => headless = True
    raw = os.getenv("SPRINGAHEAD_HEADLESS", "1").strip().lower()
    headless = raw not in ("0", "false", "no", "off")

    worked_days = fetch_worked_days(creds, headless=headless)

    if not worked_days:
        print("No worked days with hours > 0 found on this timecard.")
        return

    print("\nWorked days on current timecard (hours > 0):")
    for entry in worked_days:
        print(f"- {entry['date']} | {entry['hours']} hours | {entry['project']} ({entry['type']})")

    data = {"entries": worked_days}
    OUTPUT_JSON.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"\nSaved data to {OUTPUT_JSON.resolve()}")



if __name__ == "__main__":
    main()
