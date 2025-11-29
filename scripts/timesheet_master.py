import os
import sys
from pathlib import Path
import traceback
import pywintypes
from springahead_step2_invoice import main as step2_main

import springahead_step1_fetch as step1
import springahead_step2_invoice as step2

def get_app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

# Make stdout flush on every newline so Gooey shows output immediately
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(line_buffering=True)

def log(msg: str = ""):
    """
    Send output to Gooey in an encoding-safe way.

    - Forces message to plain ASCII.
    - Replaces any non-ASCII characters with '?' so Gooey's
      stdout reader doesn't die on weird bytes inside the EXE.
    """
    safe = str(msg).encode("ascii", "replace").decode("ascii")
    print(safe, flush=True)

def main(gui_mode=False):
    # Make sure we're running from the folder where this script lives
    base_dir = get_app_root()
    os.chdir(base_dir)

    log("======================================")
    log("  Timesheet Automation – Master Script")
    log("======================================\n")

    # ---------- STEP 1 ----------
    log("[1/2] Running Step 1 – Fetching hours from SpringAhead...")
    try:
        step1.main()
    except Exception as e:
        if gui_mode:
            # Let the GUI's outer try/except handle logging & popup
            raise
        log("\n[ERROR] Step 1 (SpringAhead fetch) failed.")
        log(f"Reason: {e}")
        traceback.print_exc()
        _pause_if_double_clicked()
        return


    json_path = base_dir / "springahead_current_week.json"
    if not json_path.exists():
        log("\n[ERROR] Step 1 finished but 'springahead_current_week.json' was not found.")
        log("Aborting before Excel step.")
        _pause_if_double_clicked()
        return

    log("\nStep 1 completed successfully.")
    log(f"  -> Data saved to: {json_path}\n")

    # ---------- STEP 2 ----------
    log("[2/2] Running Step 2 – Filling Excel invoice and exporting PDF...")
    try:
        step2_main()
        log("[INFO] Step 2 completed successfully.")
    except pywintypes.com_error as e:
        if gui_mode:
            # Let GUI show the error; don't swallow it here
            raise
        # Excel "Call was rejected by callee."
        if getattr(e, "hresult", None) == -2147418111:
            log("\n[ERROR] Step 2 (Excel/PDF) failed.")
            log("Reason: Excel rejected the automation call.")
            log(
                "\nMake sure ALL Excel windows and any EXCEL.EXE processes are closed, "
                "then try again."
            )
            _pause_if_double_clicked()
            return
        else:
            # For any other COM error, keep the original traceback
            log("\n[ERROR] Step 2 (Excel/PDF) failed with an unexpected COM error.")
            log(e)
            traceback.print_exc()
            _pause_if_double_clicked()
            return
    except Exception as e:
        if gui_mode:
            # This includes our RuntimeError("Consultant name is missing...")
            raise
        # Non-COM errors in Step 2 (CLI mode only)
        log("\n[ERROR] Step 2 (Excel/PDF) failed.")
        log(f"Reason: {e}")
        traceback.print_exc()
        _pause_if_double_clicked()
        return

    log("\nAll steps completed successfully ")
    log("You should now have:")
    log(f"  - JSON file: {json_path.name}")
    log("  - A new PDF invoice in this same folder (named like 'J. Pepin INV (...).pdf').")

    if not gui_mode:
        _pause_if_double_clicked()




def _pause_if_double_clicked():
    """
    If the script was launched by double-click (console with stdin attached),
    pause at the end so the window doesn't disappear instantly.

    In the frozen EXE (Gooey), we NEVER pause.
    """
    if getattr(sys, "frozen", False):
        return

    try:
        if sys.stdin and sys.stdin.isatty():
            input("\nPress Enter to exit...")
    except Exception:
        pass


if __name__ == "__main__":
    main()
