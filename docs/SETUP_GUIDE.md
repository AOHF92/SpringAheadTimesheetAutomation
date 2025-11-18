# 🛠️ SpringAhead Timesheet Automation – Setup Guide

This guide walks you through installing everything required to run the SpringAhead automation scripts on Windows 10 or Windows 11.

---

## 🧰 1. Requirements

### Operating System
- Windows 10 or Windows 11

### Applications
- **Microsoft Excel (desktop)**  
  Required for filling the invoice template and exporting the PDF.

### Python
- **Python 3.10 or newer**

If you don’t have Python:

1. Go to the official download page: https://www.python.org/downloads/
2. Download the latest Python 3.x for Windows.
3. **Important:** During install, check the box:
   - `Add Python to PATH`

### SpringAhead Account
- A valid SpringAhead account with access to your timecard.

## 2. Get the project files

You can get the project in one of two ways:

### Option A – Download as ZIP (simple)

1. Go to the GitHub repository page.
2. Click the green **Code** button.
3. Click **Download ZIP**.
4. Extract it somewhere easy, for example:
   - `C:\SpringAheadAutomation`

You should end up with something like:

```text
C:\SpringAheadAutomation\
  springahead_step1_fetch.py
  springahead_step2_invoice.py
  timesheet_master.py
  INVOICE (Template).xls
  requirements.txt      (if included)
```
### Option B – Using Git (for Git users)
If you have Git installed, you can clone instead:

```
git clone https://github.com/AOHF92/SpringAheadTimesheetAutomation.git
cd springahead-timesheet-automation
```
## 3. Install Python dependencies
The simplest way is to install everything globally on your system.   
Open Command Prompt or PowerShell, go to the project folder, for example:
```bash
cd C:\SpringAheadAutomation
```

### 3.1 Using requirements.txt (recommended)
Run the following:
```bash
pip install -r requirements.txt
```

This will install all required packages, such as:  

- playwright
- python-dotenv
- pywin32

Then install the Playwright browsers:
```bash
playwright install
```

### 3.2 Without requirements.txt:
You can also install the dependecies manually as an alternative:
```
pip install playwright python-dotenv pywin32
playwright install
```

## 4 Invoice Template
The project uses an Excel file named:
```text
INVOICE (Template).xls
```
- A sanitized template is included in the repo.
- Keep this file in the **same folder** as the python scripts
Note:
> If you customize it, keep the same filename and general layout expected by the script
(consultant name cell, invoice number cell, period cell, time entry rows, etc.).

## 5. Configure credentials (MyCreds.env)
The script uses a local file named ```MyCreds.env``` to store your SpringAhead login info.

### Option A - Let the script create it
If ```MyCreds.env``` is missing or incomplete:

- The script will ask you for: 

 - Company
 - Login Name
 - Password

- It will then offer to save those values into ```MyCreds.env``` for future runs.

Note: 
> ```MyCreds.env``` is a plain-text file. Keep it private and do not share it.

## 6. Running the automation

The easiest way to use the project is via the **master script**.

### 6.1 Run from the terminal
In Command Prompt or PowerShell, from the project folder:
```bash
python timesheet_master.py
```
What it will do:

1. Use Playwright to log into SpringAhead
2. Scrape your current week’s worked days
3. Save them into springahead_current_week.json
4. Open the Excel template
5. Fill in the invoice with your hours
6. Export a PDF invoice

At the end, you should see a new PDF invoice in the same folder.

### 6.2 Run by double-click (Windows Explorer)

You can also:

1. Open the project folder in Explorer
2. Double-click ```timesheet_master.py```

A console window will open and run the same process.
The script is written to pause at the end so you can read any messages before the window closes.

---

## Optional: Using a virtual environment

This section is optional.
You **do not** need a virtual environment to use this project.

A virtual environment is useful if:

- You work on multiple Python projects, or
- You want to keep this project’s dependencies isolated.

### 1. Create a virtual environment
From the project folder:
```bash
python -m venv .venv
```
### 2. Activate it (Windows)
```bash
.\.venv\Scripts\activate
```
You should see something like:
```bash
(.venv) C:\SpringAheadAutomation>
```
### 3. Install dependencies inside the venv
```bash
pip install -r requirements.txt
playwright install
```
or:
```bash
pip install playwright python-dotenv pywin32
playwright install
```
Then run the script:
```bash
python timesheet_master.py
```
### 4. Deactive the virtual environment:
When you're done:
```bash
deactivate
```
---
