import pandas as pd
import os
import time
from datetime import datetime

# === CONFIGURATION ===
CONFIG = {
    "input_file": "input_data.xlsx",                # Excel file to read
    "input_sheet": "Sheet1",                        # Sheet to read from (optional)
    "output_file": "filtered_output.xlsx",          # Excel file to write
    "columns_to_extract": [                         # List of column names to retain
        "App ID", "Application Name", "Severity", "Status"
    ],
    "drop_empty_rows": True                         # Optionally drop rows with all values as NaN
}

def print_header(title):
    print(f"\n{'=' * 60}\n🔷 {title}\n{'=' * 60}")

def extract_columns(config):
    start_time = time.time()
    print_header("STEP 1: STARTING EXTRACTION PROCESS")

    input_file = config["input_file"]
    print(f"📁 Checking if file exists: {input_file}")
    if not os.path.exists(input_file):
        print(f"❌ ERROR: File not found: {input_file}")
        return

    try:
        print(f"📄 Reading Excel file: {input_file}, sheet: {config.get('input_sheet', 0)}")
        df = pd.read_excel(input_file, sheet_name=config.get("input_sheet", 0))
        print(f"✅ SUCCESS: Loaded sheet with {df.shape[0]} rows and {df.shape[1]} columns.")
    except Exception as e:
        print(f"❌ ERROR: Failed to read Excel file. Details: {e}")
        return

    print_header("STEP 2: VALIDATING COLUMNS")

    requested_cols = config["columns_to_extract"]
    print(f"📌 Columns requested: {requested_cols}")

    missing_cols = [col for col in requested_cols if col not in df.columns]
    if missing_cols:
        print(f"⚠️ WARNING: These columns are missing and will be skipped: {missing_cols}")

    selected_cols = [col for col in requested_cols if col in df.columns]
    print(f"✅ Columns to extract: {selected_cols}")

    if not selected_cols:
        print("❌ ERROR: None of the requested columns were found. Aborting.")
        return

    df_filtered = df[selected_cols]
    original_row_count = df_filtered.shape[0]

    if config.get("drop_empty_rows", False):
        df_filtered = df_filtered.dropna(how="all")
        new_row_count = df_filtered.shape[0]
        print(f"🧹 Dropped rows with all empty values. Rows reduced from {original_row_count} to {new_row_count}.")
    else:
        print("ℹ️ Skipping drop of empty rows as per config.")

    print_header("STEP 3: WRITING OUTPUT FILE")

    try:
        df_filtered.to_excel(config["output_file"], index=False)
        print(f"✅ SUCCESS: Output written to '{config['output_file']}' with {df_filtered.shape[0]} rows and {df_filtered.shape[1]} columns.")
    except Exception as e:
        print(f"❌ ERROR: Failed to write output file. Details: {e}")
        return

    end_time = time.time()
    duration = end_time - start_time

    print_header("✅ EXTRACTION COMPLETED")
    print(f"🕒 Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"⏱️ Total Execution Time: {duration:.2f} seconds\n")

# === Run Script ===
if __name__ == "__main__":
    extract_columns(CONFIG)
