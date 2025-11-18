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

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from dotenv import load_dotenv


LOGIN_URL = (
    "https://my.springahead.com/go/Account/Logon"
    "?ReturnUrl=%2Fvt%2Fgo%3FHome%26tokenid%3Dvte"
)

ENV_PATH = Path("MyCreds.env")
OUTPUT_JSON = Path("springahead_current_week.json")


def load_credentials():
    """
    Try to load credentials from .env.
    If missing/incomplete, prompt the user and optionally write .env.
    """
    creds = {
        "company": None,
        "username": None,
        "password": None,
    }

    env_ok = False
    if ENV_PATH.exists():
        load_dotenv(dotenv_path=ENV_PATH)
        creds["company"] = os.getenv("SPRINGAHEAD_COMPANY") or ""
        creds["username"] = os.getenv("SPRINGAHEAD_USERNAME") or ""
        creds["password"] = os.getenv("SPRINGAHEAD_PASSWORD") or ""

        if all(creds.values()):
            env_ok = True

    if env_ok:
        print("Loaded SpringAhead credentials from MyCreds.env.")
        return creds

    # Otherwise, ask interactively
    print("No complete .env found. Please enter your SpringAhead credentials.")
    company = input("Company (e.g., MetroIT): ").strip()
    username = input("Login Name: ").strip()
    password = getpass("Password: ").strip()

    creds["company"] = company
    creds["username"] = username
    creds["password"] = password

    # Offer to create/update .env
    save = input("Save these credentials to .env for next time? [y/N]: ").strip().lower()
    if save == "y":
        with ENV_PATH.open("w", encoding="utf-8") as f:
            f.write(f"SPRINGAHEAD_COMPANY={company}\n")
            f.write(f"SPRINGAHEAD_USERNAME={username}\n")
            f.write(f"SPRINGAHEAD_PASSWORD={password}\n")
        print("Saved credentials to MyCreds.env (plain text – keep this file private).")

    return creds


def fetch_worked_days(creds, headless=False):
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
    worked_days = fetch_worked_days(creds, headless=False)

    if not worked_days:
        print("No worked days with hours > 0 found on this timecard.")
        return

    print("\nWorked days on current timecard (hours > 0):")
    for entry in worked_days:
        print(f"- {entry['date']} | {entry['hours']} hours | {entry['project']} ({entry['type']})")

    # Save to JSON for the Excel step later
    data = {
        "entries": worked_days,
    }
    OUTPUT_JSON.write_text(json.dumps(data, indent=2), encoding="utf-8")
    print(f"\nSaved data to {OUTPUT_JSON.resolve()}")


if __name__ == "__main__":
    main()
