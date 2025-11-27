# 🛠️ SpringAhead Timesheet Generator – Setup & Usage Guide (v2.0.0)

This guide explains how to use the new GUI version, the standalone Windows executable, and how to run the project from source for development.

---

# 📦 1. Overview

This project automates:

1. **Scraping worked hours from SpringAhead** (via Playwright)
2. **Generating a professional PDF invoice** (via Excel Automation)

Version **2.0.0** introduces:

- A full **Graphical User Interface (GUI)**
- A standalone **Windows executable** (`InvoiceGenerator.exe`)
- Better logging, safer file handling, and improved stability

---

# 🚀 2. Quick Start (Recommended for Regular Users)

This is the easiest way to use the tool — **no Python required**.

### ✅ Requirements
- Windows 10 or Windows 11
- Microsoft Excel (desktop version)
- A SpringAhead account with timecard access

### 📁 Files to Download
From the latest GitHub Release:

- **`InvoiceGenerator.exe`**
- **`INVOICE (Template).xls`**

> ⚠ Place both files in the **same folder**.

---

## ▶️ How to Use the EXE

1. Create a folder, e.g.  
   `D:\SpringAheadTimeSheetGenerator\`

2. Put these files inside it:
- ``InvoiceGenerator.exe``
- ``INVOICE (Template).xls``

3. Double-click **InvoiceGenerator.exe**.

4. On first run, enter your:
- Company  
- Login Name  
- Password  

(Optional) Check **“Save credentials to MyCreds.env”**  
→ the app will create `MyCreds.env` automatically.

5. Choose a mode:
- **Full pipeline (recommended)**
- **Step 1 only** – scrape worked hours (JSON)
- **Step 2 only** – generate invoice from JSON

6. Click **Start**.

7. When the process completes, you will find:
- `springahead_current_week.json`
- A new PDF invoice like:  
  `J. Pepin INV (MM-DD-YYYY).pdf`

All generated files appear in the **same folder as the EXE**.

---

# 🛠️ 3. Developer Setup (Run From Source)

If you want to modify or run the project via Python, follow this section.

---

# 📥 3.1 Download the Project

### Option A — Download ZIP
1. Go to the GitHub repo → **Code** → **Download ZIP**
2. Extract anywhere, e.g.  
`C:\SpringAheadAutomation\`

### Option B — Clone with Git
```bash
git clone https://github.com/AOHF92/SpringAheadTimesheetAutomation.git
cd SpringAheadTimesheetAutomation
```
### Your folder should contain:
```java
springahead_gui.py
springahead_step1_fetch.py
springahead_step2_invoice.py
timesheet_master.py
INVOICE (Template).xls
requirements.txt
```
---

## 🐍 3.2 Install Python

Download Python from:

https://www.python.org/downloads/

During installation, make sure to check:

- ✅ **Add Python to PATH**

---

## 📦 3.3 Install Dependencies

Open **Command Prompt** or **PowerShell** inside the project directory.

### Recommended (using `requirements.txt`)

```bash
pip install -r requirements.txt
python -m playwright install chromium
```
Manual install (alternative)
```bash
pip install playwright python-dotenv pywin32 gooey
python -m playwright install chromium
```

# 🔐 4. Credentials (`MyCreds.env`)

Your SpringAhead credentials are stored in a local file called **MyCreds.env**.

### 📄 File Format

```env
SPRINGAHEAD_COMPANY=YourCompany
SPRINGAHEAD_USERNAME=yourlogin
SPRINGAHEAD_PASSWORD=yourpassword
```
### How to Create It

#### ✅ Option A – Automatic (Recommended)

If `MyCreds.env` is missing, the GUI will:

1. Prompt you for your SpringAhead login details  
2. Offer to save them into `MyCreds.env` for future runs

The file will be created in the **same folder** as:

- `InvoiceGenerator.exe` (EXE users), or  
- the Python scripts (source/development users)

---

#### ✏️ Option B – Manual Creation

Create a new file named `MyCreds.env` in the same directory as the EXE or scripts.

Use this format:
```yaml
SPRINGAHEAD_COMPANY=YourCompany
SPRINGAHEAD_USERNAME=yourlogin
SPRINGAHEAD_PASSWORD=yourpassword
```
> ⚠ **Warning:**  
> `MyCreds.env` is plain text.  
> Do **not** upload it to GitHub or share it with others.

---

# 📄 5. Excel Template — `INVOICE (Template).xls`

This file is required for generating your invoice PDF.

### Placement Requirements

It **must** be located in the same folder as:

- `InvoiceGenerator.exe` (when using the EXE)
- The Python scripts (when running from source)

Your folder should look like:
```yaml
InvoiceGenerator.exe
INVOICE (Template).xls
MyCreds.env (optional; created automatically)
```
### Notes on Customizing the Template

- You may modify the visual layout  
- However, keep the same filename  
- And keep the fields/cells the script expects  
  (e.g., consultant name, invoice number, billing period, rows for hours)

Changing those requires updating the Python code.

---

# ▶️ 6. Running the Program (From Source)

If you are developing or modifying the project, you can run it directly in Python.

---

## 🖥️ 6.1 Using the GUI (Recommended)

From the project directory:
```python
python springahead_gui.py
```
Steps:

1. Enter (or confirm) your credentials  
2. Choose the mode:
   - **Full pipeline**
   - **Step 1 only** (scrape → JSON)
   - **Step 2 only** (JSON → PDF)
3. Click **Start**
4. Watch progress appear in the GUI output panel

This provides the same workflow as the standalone EXE.

---

## 🤖 6.2 Running the Master Script Directly

The old CLI behavior still works:
```python
python timesheet_master.py
```
This will:

1. Start Playwright  
2. Log into SpringAhead  
3. Scrape your worked hours  
4. Store them in `springahead_current_week.json`  
5. Fill the Excel invoice template  
6. Export the invoice as a PDF

Useful for debugging or terminal-based workflows.

---

# 🔧 7. Optional: Virtual Environment (For Developers)

A virtual environment isolates dependencies and prevents conflicts with other Python projects.

---

## 7.1 Create the Environment

```shell
python -m venv .venv
```

## 7.2 Activate the Environment (Windows)

```sql
..venv\Scripts\activate
```

Your prompt should now start with:
``` (.venv) ```

## 7.3 Install Dependencies Inside the venv

With the virtual environment activated, install all necessary packages:

```bash
pip install -r requirements.txt
python -m playwright install chromium
```

Or install them manually:

```bash
pip install playwright python-dotenv pywin32 gooey
python -m playwright install chromium
```

This ensures the GUI, Playwright browser automation, Excel automation, and environment variable loader all work correctly.

---

## 7.4 Run the GUI

Once dependencies are installed, launch the application:

```python
python springahead_gui.py
```

You will see the full graphical interface where you can:

- Enter your SpringAhead credentials  
- Select the pipeline mode  
- Generate invoices  
- Save credentials to `MyCreds.env` (optional)  

This is the same interface bundled into the EXE.

---

## 7.5 Deactivate the venv

When you are finished working in the environment:
```bash
deactivate
```

Your terminal prompt will return to normal, and your global Python installation will no longer be affected by this project’s dependencies.

---

# 🏁 8. Output Files

After running the automation (either via EXE or source), the following files will appear in the same folder:

```yaml
MyCreds.env # created only if you choose to save credentials
springahead_current_week.json # raw scraped worked-hours data
YourName INV (MM-DD-YYYY).pdf # the generated invoice
```

All outputs are always written to the folder containing the EXE or the Python scripts.

---
