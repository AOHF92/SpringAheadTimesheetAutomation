from gooey import Gooey, GooeyParser
from pathlib import Path
import os, sys, io

import timesheet_master as tm
import springahead_step1_fetch as step1
import springahead_step2_invoice as step2

def get_app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

# Force a predictable stdout encoding for Gooey/PyInstaller combo
if sys.stdout and not sys.stdout.encoding:
    sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding="utf-8", errors="replace", line_buffering=True)

@Gooey(
    program_name="SpringAhead Invoice Generator",
    default_size=(800, 600),
    required_cols=1,
    optional_cols=1,
    clear_before_run=True,
    show_success_modal=True,
)
def main():
    # Make sure we run from this script's folder (like your master script does)
    base_dir = get_app_root()
    os.chdir(base_dir)

    parser = GooeyParser(
        description=(
            "Run this to fetch your current week and fill the Excel invoice.\n\n"
            "You can either run everything, or step by step.\n"
            "If you already have a .env with your credentials, just leave the credential fields blank."
        )
    )

    # --- What do you want to run? ---
    parser.add_argument(
        "mode",
        choices=["Full pipeline (Step 1 + Step 2)",
                 "Step 1 only (Fetch from SpringAhead)",
                 "Step 2 only (Invoice from JSON)"],
        default="Full pipeline (Step 1 + Step 2)",
        help="Choose which part of the automation to run.",
    )
    
        # --- Browser visibility ---
    parser.add_argument(
        "--show-browser",
        action="store_true",
        default=False,  # unchecked by default => headless
        help="Show the browser window while fetching (disables headless mode for Step 1).",
    )

    # --- Credentials section (optional override) ---
    creds_group = parser.add_argument_group(
        "SpringAhead Credentials (optional override)",
        "If left blank, the script will use .env variables\n"
        "or fall back to the original console prompts."
    )

    creds_group.add_argument(
        "--company",
        metavar="Company",
        help="SpringAhead company (e.g., MetroIT). Leave blank to use .env / default behavior.",
    )
    creds_group.add_argument(
        "--username",
        metavar="Login Name",
        help="SpringAhead login name. Leave blank to use .env / default behavior.",
    )
    creds_group.add_argument(
        "--password",
        metavar="Password",
        help="SpringAhead password. Leave blank to use .env / default behavior.",
        widget="PasswordField",
    )
    creds_group.add_argument(
        "--save-env",
        action="store_true",
        default=False,
        help="Save these credentials for future runs?(A .env file will be created.)",
    )

    # (Optional) name field – this just helps avoid prompts if your template is blank
    name_group = parser.add_argument_group(
        "Invoice Options",
        "Your full name is normally pulled from the Excel file if already present.\n"
        "If your invoice template is blank, be sure to write your name here atleast once."
    )
    name_group.add_argument(
        "--full-name",
        metavar="Consultant full name",
        help="Optional: full name to pre-fill in the invoice template (e.g., John Doe).",
    )

    args = parser.parse_args()
    
    # --- Headless toggle for Step 1 via env var ---
    #
    # springahead_step1_fetch.main() will read SPRINGAHEAD_HEADLESS and
    # default to headless = True when it's not set / not "0"/"false"/"no"/"off".
    #
    if args.show_browser:
        # User wants to see the browser -> headed mode
        os.environ["SPRINGAHEAD_HEADLESS"] = "0"
    else:
        # Default: run headless
        os.environ["SPRINGAHEAD_HEADLESS"] = "1"
    # --- Apply credential overrides via environment variables ---
    #
    # springahead_step1_fetch.py already reads:
    #   SPRINGAHEAD_COMPANY / SPRINGAHEAD_USERNAME / SPRINGAHEAD_PASSWORD
    # from the environment and/or .env.
    #
    # Here we set the env vars for this run only, without touching the .env file.
    #
    if args.company:
        os.environ["SPRINGAHEAD_COMPANY"] = args.company
    if args.username:
        os.environ["SPRINGAHEAD_USERNAME"] = args.username
    if args.password:
        os.environ["SPRINGAHEAD_PASSWORD"] = args.password
        
    # --- Optionally persist credentials into .env ---
    if args.save_env:
        env_path = base_dir / "MyCreds.env"

        company = os.environ.get("SPRINGAHEAD_COMPANY", "")
        username = os.environ.get("SPRINGAHEAD_USERNAME", "")
        password = os.environ.get("SPRINGAHEAD_PASSWORD", "")

        # Simple overwrite – if you want, we can later implement a "merge" instead.
        try:
            with env_path.open("w", encoding="utf-8") as f:
                if company:
                    f.write(f'SPRINGAHEAD_COMPANY="{company}"\n')
                if username:
                    f.write(f'SPRINGAHEAD_USERNAME="{username}"\n')
                if password:
                    f.write(f'SPRINGAHEAD_PASSWORD="{password}"\n')

            print(f"Saved credentials to {env_path}")
        except Exception as e:
            print(f"[WARN] Could not write .env file: {e}")

    # Optional: make the full name visible to step 2 by env var
    # (we'll extend step 2 later to check this instead of calling input())
    if getattr(args, "full_name", None):
        os.environ["SPRINGAHEAD_FULL_NAME"] = args.full_name

    # --- Dispatch based on selected mode ---
    mode = args.mode

    try:
        if mode.startswith("Full pipeline"):
            # Use your existing orchestration logic
            tm.main()

        elif mode.startswith("Step 1 only"):
            step1.main()

        elif mode.startswith("Step 2 only"):
            step2.main()

        else:
            raise ValueError(f"Unknown mode: {mode}")

    except Exception as e:
        # Let Gooey show an error in its console
        print("\n[ERROR] Something went wrong during execution.")
        print(f"Reason: {e}")
        # If you want full traceback inside Gooey's console:
        import traceback
        traceback.print_exc()
        # Non-zero exit so Gooey shows failure state
        sys.exit(1)


if __name__ == "__main__":
    main()