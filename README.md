# SpringAhead Timesheet Automation

Automates the timesheet generation process of SpringAhead workflow:

1. Logs into **SpringAhead** using Playwright.
2. Scrapes the **current week’s worked days** (only days with hours > 0).
3. Fills an **Excel invoice template** with those hours.
4. Exports a **PDF invoice** ready for manual signature.

---

## Features

- ✅ **Step 1 – Fetch SpringAhead hours**  
  - Uses Playwright to log into SpringAhead.
  - Switches to *List* view and scrapes the current timecard.
  - Saves the result as `springahead_current_week.json`.

- ✅ **Step 2 – Fill Excel invoice & export PDF**  
  - Opens `INVOICE (Template).xls` via the Excel COM API (`pywin32`).
  - Automatically increments the invoice number.
  - Computes standard morning/afternoon time blocks based on total hours.
  - Fills the invoice lines (date, from, to, task).
  - Exports a PDF invoice (filename includes a short consultant name + period).

- ✅ **Master script – One-shot run**  
  - Runs Step 1 and Step 2 in sequence.
  - Shows clear console messages and friendly error handling.
  - Pauses at the end when launched by double-click so the console doesn’t vanish.

- 🔜 **Planned**  
  - Step 3: Email automation (e.g., generate an Outlook email with the PDF attached).

---

## Requirements

- OS: Windows 10 or Windows 11
- Python: 3.10+ (tested on Windows)
- Excel: Microsoft Excel (desktop, with COM automation enabled)
- SpringAhead account with access to the timecard

### Python dependencies

Install these with pip:

- playwright
- python-dotenv
- pywin32

---

## ⚙️ Installation & Setup Guide

For the full detailed setup steps, see:  

➡️ [📦 Full Installation & Setup Guide](docs/SETUP_GUIDE.md)

---
## Usage

### Option A – Run everything via the master script

From a terminal:
```bash
python timesheet_master.py
```

What it does:

- Runs Step 1 – logs into SpringAhead, scrapes the current week, and writes ```springahead_current_week.json```.
- Runs Step 2 – opens the Excel template, fills the invoice, and exports a PDF.

At the end you should have:

- ```springahead_current_week.json``` with your worked days.
- A new PDF invoice in the same folder, named like:

```text
J. Pepin INV (11 - 1 al 15 - 2025).pdf
```
Note:
> The short consultant name (e.g. J. Pepin) is derived from the full name stored in cell B6 of the template or entered on first run.

If you prefer, you can also double-click ```timesheet_master.py``` in Explorer.
The script will pause at the end so you can read any messages before the window closes.

### Option B – Run steps individually

Step 1 – Fetch SpringAhead hours
```bash
python springahead_step1_fetch.py
```
This will:

- Prompt for credentials (or use ```MyCreds.env```).
- Log into SpringAhead.
- Click “Add Time” and switch to List view.
- Scrape all time rows with hours > 0.
- Print them to the console and write:
```
springahead_current_week.json
```

Step 2 – Fill Excel invoice and export PDF
```
python springahead_step2_invoice.py
```
This expects:

- ```springahead_current_week.json``` in the same folder.
- ```INVOICE (Template).xls``` in the same folder.

It will:

- Load JSON entries and compute the period string (e.g. 11 - 1 al 15 - 2025).
- Increment the invoice number.
- Fill time entries in the invoice.
- Save the workbook and export a PDF invoice.

---

## File glossary
- ```timesheet_master.py```
Orchestrator script. Runs Step 1 and Step 2 in sequence, with error handling and “press Enter to exit…” behavior when double-clicked.
- ```springahead_step1_fetch.py``` 
Playwright scraper:  
  - Loads credentials from ```MyCreds.env``` (or interactively).
  - Logs into SpringAhead.
  - Switches to List view.
  - Scrapes worked days from the current timecard.
  - Saves them to ```springahead_current_week.json```.
- ```springahead_step2_invoice.py```
Excel automation:   
   - Reads ```springahead_current_week.json```.
   - Detects the invoice period (first or second half of the month).
   - Calculates morning/afternoon time blocks based on total hours.
   - Fills the invoice template and exports a PDF.
- ```INVOICE (Template).xls```
Local Excel Invoice Template
Contains your layout, rates, and formulas.
- ```MyCreds.env```
Local credentials file containing SPRINGAHEAD_COMPANY, SPRINGAHEAD_USERNAME, and SPRINGAHEAD_PASSWORD.
Plain-text. Keep it private
- ```springahead_current_week.json```
Generated Json File with structure:
```json
{
  "entries": [
    {
      "date": "11/02/2025",
      "hours": 8.0,
      "project": "Some Project Name",
      "type": "Regular"
    }
  ]
}
```
