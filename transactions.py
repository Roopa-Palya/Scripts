import pandas as pd
import os
import time
from datetime import datetime

# === CONFIGURATION SECTION ===
CONFIG = {
    "input_file": "input_data.xlsx",           # Input Excel file
    "input_sheet": "Sheet1",                   # Sheet name in Excel
    "output_file": "final_output.xlsx",        # Output Excel file

    "columns_to_extract": [                    # Columns to keep from input
        "App ID", "Application Name", "Scan Date", "Age (days)"
    ],

    "columns_to_add": {                        # Additional static or blank columns
        "Reviewed By": "Security Team",
        "Status": ""
    },

    "derived_column": {                        # Fill existing column with calculated value
        "column_name": "Age (days)",           # Existing column to be filled
        "based_on_column": "Scan Date"         # Use this date to calculate age
    },

    "drop_empty_rows": True                    # Drop rows with all empty values
}

# === HELPER FUNCTIONS ===
def print_header(title):
    print(f"\n{'=' * 60}\n🔷 {title}\n{'=' * 60}")

# === MAIN PROCESSING FUNCTION ===
def process_excel(config):
    start_time = time.time()

    # STEP 1: LOAD EXCEL FILE
    print_header("STEP 1: LOADING INPUT FILE")
    if not os.path.exists(config["input_file"]):
        print(f"❌ ERROR: File not found: {config['input_file']}")
        return

    try:
        df = pd.read_excel(config["input_file"], sheet_name=config.get("input_sheet", 0))
        print(f"✅ Loaded '{config['input_file']}' with {df.shape[0]} rows and {df.shape[1]} columns.")
    except Exception as e:
        print(f"❌ Failed to read Excel file: {e}")
        return

    # STEP 2: SELECT DESIRED COLUMNS
    print_header("STEP 2: SELECTING COLUMNS")
    requested_cols = config["columns_to_extract"]
    print(f"📌 Requested columns: {requested_cols}")
    missing = [col for col in requested_cols if col not in df.columns]
    if missing:
        print(f"⚠️ WARNING: These columns are missing and will be skipped: {missing}")
    df = df[[col for col in requested_cols if col in df.columns]]

    if config.get("drop_empty_rows", False):
        before = df.shape[0]
        df = df.dropna(how='all')
        print(f"🧹 Dropped {before - df.shape[0]} completely empty rows.")

    # STEP 3: ADD NEW COLUMNS WITH STATIC OR BLANK VALUES
    print_header("STEP 3: ADDING STATIC/BLANK COLUMNS")
    for col, value in config["columns_to_add"].items():
        df[col] = value
        print(f"➕ Added column '{col}' with value: '{value}'")

    # STEP 4: UPDATE EXISTING COLUMN WITH CALCULATED AGE IN DAYS
    print_header("STEP 4: CALCULATING AGE IN DAYS")
    derived = config["derived_column"]
    base_col = derived["based_on_column"]
    target_col = derived["column_name"]

    if base_col in df.columns and target_col in df.columns:
        today = datetime.today().date()
        print(f"📆 Calculating number of days since '{base_col}' to today ({today})...")

        df[base_col] = pd.to_datetime(df[base_col], errors='coerce').dt.date
        df[target_col] = df[base_col].apply(
            lambda d: (today - d).days if pd.notnull(d) else ""
        )

        print(f"✅ Column '{target_col}' updated successfully.")
    else:
        print(f"⚠️ Either '{base_col}' or '{target_col}' is missing. Skipping age calculation.")

    # STEP 5: WRITE OUTPUT FILE
    print_header("STEP 5: WRITING OUTPUT FILE")
    try:
        df.to_excel(config["output_file"], index=False)
        print(f"✅ Output written to '{config['output_file']}' with {df.shape[0]} rows and {df.shape[1]} columns.")
    except Exception as e:
        print(f"❌ Failed to write output file: {e}")
        return

    # STEP 6: EXECUTION SUMMARY
    end_time = time.time()
    print_header("✅ PROCESS COMPLETED")
    print(f"🕒 Finished at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"⏱️ Total Execution Time: {end_time - start_time:.2f} seconds\n")

# === ENTRY POINT ===
if __name__ == "__main__":
    process_excel(CONFIG)
