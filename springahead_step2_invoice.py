"""
Step 2 – Fill Excel invoice from SpringAhead JSON and export to PDF.

Requirements:
    pip install pywin32

Files expected in the same folder as this script:
    - springahead_current_week.json   (output of Step 1)
    - INVOICE (Template).xls          (your invoice template)
"""

import os
import json
import calendar
import re
from datetime import datetime, timedelta

import win32com.client as win32


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(SCRIPT_DIR, "springahead_current_week.json")
TEMPLATE_PATH = os.path.join(SCRIPT_DIR, "INVOICE (Template).xls")


# ---------- Helpers ----------

def load_entries_from_json(path):
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    entries = data.get("entries", [])

    def parse_date(e):
        return datetime.strptime(e["date"], "%m/%d/%Y")

    return sorted(entries, key=parse_date)


def detect_period_string(entries):
    """
    Build period string like '11 - 1 al 15 - 2025'
    based on month/year of the entries and which half of the month.
    """
    if not entries:
        raise ValueError("No entries in JSON; cannot compute period.")

    dt0 = datetime.strptime(entries[0]["date"], "%m/%d/%Y")
    month = dt0.month
    year = dt0.year

    max_day = max(datetime.strptime(e["date"], "%m/%d/%Y").day for e in entries)
    if max_day <= 15:
        start_day, end_day = 1, 15
    else:
        start_day, end_day = 16, calendar.monthrange(year, month)[1]

    period_str = f"{month} - {start_day} al {end_day} - {year}"
    return period_str


def safe_filename(name: str) -> str:
    """Remove characters that are invalid in Windows filenames."""
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def parse_consultant_name(full_name):
    """
    Parse hispanic full names

    Rules:
      - First word = first name
      - If second word is an initial (e.g. 'O.') → treat it as middle name
      - First last name = next available word after middle names
    """
    full_name = full_name.strip()
    parts = full_name.split()
    if len(parts) < 1:
        raise ValueError("Name cannot be empty.")

    first_name = parts[0]

    # Default assumption
    first_last = None

    # Detect middle-name initial (e.g. O.)
    if len(parts) >= 3 and re.fullmatch(r"[A-Za-z]\.", parts[1]):
        # Skip the initial → last name starts at parts[2]
        first_last = parts[2]
    else:
        # Otherwise, second word IS the first last name
        if len(parts) >= 2:
            first_last = parts[1]
        else:
            first_last = first_name  # fallback only if no last name exists

    initial = first_name[0].upper()
    short_name = f"{initial}. {first_last}"
    return full_name, short_name


def compute_time_blocks(total_hours):
    """
    Given total hours from SpringAhead (e.g. 8.00, 8.25, 7.5),
    return:
      - morning_from, morning_to
      - afternoon_from, afternoon_to

    Base:
      morning:  7:00 AM -> 11:00 AM (4h)
      afternoon: 12:00 PM -> 4:00 PM (4h)
    Overtime/undertime adjusts only the afternoon 'To' time in 15-minute increments.
    """
    base_hours = 8.0
    dummy_date = datetime(2000, 1, 1)
    morning_from = dummy_date.replace(hour=7, minute=0)
    morning_to = dummy_date.replace(hour=11, minute=0)
    afternoon_from = dummy_date.replace(hour=12, minute=0)
    afternoon_base_to = dummy_date.replace(hour=16, minute=0)

    diff_hours = float(total_hours) - base_hours
    diff_minutes = diff_hours * 60.0

    quarter = 15
    diff_minutes_rounded = int(round(diff_minutes / quarter) * quarter)

    afternoon_to = afternoon_base_to + timedelta(minutes=diff_minutes_rounded)

    def fmt(dt):
        return dt.strftime("%I:%M %p").lstrip("0")  # e.g. "7:00 AM"

    return (
        fmt(morning_from),
        fmt(morning_to),
        fmt(afternoon_from),
        fmt(afternoon_to),
    )


# ---------- Main Excel logic ----------

def main():
    if not os.path.exists(JSON_PATH):
        raise FileNotFoundError(f"JSON not found: {JSON_PATH}")
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")

    entries = load_entries_from_json(JSON_PATH)
    period_str = detect_period_string(entries)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True  # set False if you want headless

    wb = excel.Workbooks.Open(TEMPLATE_PATH)
    ws = wb.Worksheets(1)  # assume first sheet is the invoice

    # ----- Consultant name from B6 (merged B6:F6) -----
    consultant_cell = ws.Cells(6, 2)  # B6
    existing_name = consultant_cell.Value
    if existing_name is not None:
        existing_name = str(existing_name).strip()
    else:
        existing_name = ""

    if existing_name:
        full_name_input = existing_name
    else:
        full_name_input = input("Enter your full name (first + one or two last names): ").strip()
        while not full_name_input:
            full_name_input = input("Name cannot be empty. Please enter full name: ").strip()
        # write it back into the template so next run doesn’t ask again
        consultant_cell.Value = full_name_input

    full_name, short_name = parse_consultant_name(full_name_input)

    try:
        # ----- Invoice Number (merged E4:F4 → anchor E4) -----
        invoice_cell = ws.Cells(4, 5)  # E4
        current_number = invoice_cell.Value
        if current_number is None:
            current_number = 0
        try:
            current_number = int(current_number)
        except Exception:
            current_number = 0
        new_number = current_number + 1
        invoice_cell.Value = new_number

        # ----- Period (merged E5:F5 → anchor E5) -----
        period_cell = ws.Cells(5, 5)  # E5
        period_cell.Value = period_str

        # OPTIONAL: write full name into Consultant cell if you want.
        # If you tell me the exact row/col for that, we can add it here.
        # For now, we leave the Consultant area as-is so the template
        # stays generic for your coworkers.

        # ----- Clear only A–D rows 9–38 -----
        first_data_row = 9
        last_data_row = 38
        for r in range(first_data_row, last_data_row + 1):
            ws.Cells(r, 1).Value = None  # A: Date
            ws.Cells(r, 2).Value = None  # B: From
            ws.Cells(r, 3).Value = None  # C: To
            ws.Cells(r, 4).Value = None  # D: Task

        # ----- Fill rows from JSON entries -----
        current_row = first_data_row

        for entry in entries:
            if current_row + 1 > last_data_row:
                print("Warning: not enough rows in template to fit all entries.")
                break

            date_str = entry["date"]     # e.g. "11/2/2025"
            hours_val = float(entry["hours"])
            dt = datetime.strptime(date_str, "%m/%d/%Y")

            m_from, m_to, a_from, a_to = compute_time_blocks(hours_val)

            # Morning row
            ws.Cells(current_row, 1).Value = dt
            ws.Cells(current_row, 2).Value = m_from
            ws.Cells(current_row, 3).Value = m_to
            ws.Cells(current_row, 4).Value = "Remote IT Support"

            # Afternoon row
            ws.Cells(current_row + 1, 1).Value = dt
            ws.Cells(current_row + 1, 2).Value = a_from
            ws.Cells(current_row + 1, 3).Value = a_to
            ws.Cells(current_row + 1, 4).Value = "Remote IT Support"

            current_row += 2

        # ----- Export to PDF -----
        pdf_filename = safe_filename(f"{short_name} INV ({period_str}).pdf")
        pdf_path = os.path.join(SCRIPT_DIR, pdf_filename)

        wb.Save()

        try:
            ws.ExportAsFixedFormat(
                Type=0,  # PDF
                Filename=pdf_path,
                Quality=0,
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            print(f"Invoice filled and exported to PDF:\n  {pdf_path}")
        except Exception as e:
            print("Export to PDF failed.")
            print(f"Target path: {pdf_path}")
            print(f"Error: {e}")
            print("Leaving Excel open so you can try exporting manually.")
            return

    finally:
        # When you're happy with it, uncomment to auto-close:
        wb.Close(SaveChanges=True)
        excel.Quit()
        pass


if __name__ == "__main__":
    main()
