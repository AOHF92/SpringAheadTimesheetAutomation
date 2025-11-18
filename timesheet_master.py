import os
import sys
from pathlib import Path
import traceback
import pywintypes
from springahead_step2_invoice import main as step2_main

import springahead_step1_fetch as step1
import springahead_step2_invoice as step2


def main():
    # Make sure we're running from the folder where this script lives
    base_dir = Path(__file__).resolve().parent
    os.chdir(base_dir)

    print("======================================")
    print("  Timesheet Automation – Master Script")
    print("======================================\n")

    # ---------- STEP 1 ----------
    print("[1/2] Running Step 1 – Fetching hours from SpringAhead...")
    try:
        step1.main()
    except Exception as e:
        print("\n[ERROR] Step 1 (SpringAhead fetch) failed.")
        print(f"Reason: {e}")
        traceback.print_exc()
        _pause_if_double_clicked()
        return

    json_path = base_dir / "springahead_current_week.json"
    if not json_path.exists():
        print("\n[ERROR] Step 1 finished but 'springahead_current_week.json' was not found.")
        print("Aborting before Excel step.")
        _pause_if_double_clicked()
        return

    print("\nStep 1 completed successfully.")
    print(f"  -> Data saved to: {json_path}\n")

    # ---------- STEP 2 ----------
    print("[2/2] Running Step 2 – Filling Excel invoice and exporting PDF...")
    try:
        step2_main()
        print("[INFO] Step 2 completed successfully.")
    except pywintypes.com_error as e:
        # Excel "Call was rejected by callee."
        if getattr(e, "hresult", None) == -2147418111:
            print("\n[ERROR] Step 2 (Excel/PDF) failed.")
            print("Reason: Excel rejected the automation call.")
            print(
                "\nMake sure ALL Excel windows and any EXCEL.EXE processes are closed, "
                "then try again."
            )
            _pause_if_double_clicked()
            return
        else:
            # For any other COM error, keep the original traceback
            print("\n[ERROR] Step 2 (Excel/PDF) failed with an unexpected COM error.")
            print(e)
            traceback.print_exc()
            _pause_if_double_clicked()
            return
    except Exception as e:
        # Non-COM errors in Step 2
        print("\n[ERROR] Step 2 (Excel/PDF) failed.")
        print(f"Reason: {e}")
        traceback.print_exc()
        _pause_if_double_clicked()
        return

    print("\nAll steps completed successfully 🎉")
    print("You should now have:")
    print(f"  - JSON file: {json_path.name}")
    print("  - A new PDF invoice in this same folder (named like 'J. Pepin INV (...).pdf').")

    _pause_if_double_clicked()



def _pause_if_double_clicked():
    """
    If the script was launched by double-click (console with stdin attached),
    pause at the end so the window doesn't disappear instantly.
    """
    try:
        # Only pause if running in an interactive console
        if sys.stdin and sys.stdin.isatty():
            input("\nPress Enter to exit...")
    except Exception:
        # Fail silently if stdin is weird (e.g. run from IDE)
        pass


if __name__ == "__main__":
    main()
