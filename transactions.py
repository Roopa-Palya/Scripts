import pandas as pd
import os
import time
from datetime import datetime

# === CONFIGURATION ===
CONFIG = {
    "input_file": "input_data.xlsx",
    "input_sheet": "Sheet1",
    "output_file": "final_output.xlsx",

    "columns_to_extract": [
        "App ID", "Application Name", "Severity"
    ],

    "columns_to_add": {
        "Scan Date": "2025-06-08",
        "Reviewed By": "Security Team",
        "Comments": ""
    },

    "drop_empty_rows": True
}

def print_header(title):
    print(f"\n{'=' * 60}\n🔷 {title}\n{'=' * 60}")

def process_excel(config):
    start_time = time.time()
    print_header("STEP 1: LOADING INPUT FILE")

    input_file = config["input_file"]
    if not os.path.exists(input_file):
        print(f"❌ ERROR: File not found: {input_file}")
        return

    try:
        df = pd.read_excel(input_file, sheet_name=config.get("input_sheet", 0))
        print(f"✅ Loaded '{input_file}' with {df.shape[0]} rows and {df.shape[1]} columns.")
    except Exception as e:
        print(f"❌ Failed to read Excel file: {e}")
        return

    print_header("STEP 2: SELECTING COLUMNS")

    requested_cols = config["columns_to_extract"]
    print(f"📌 Requested columns: {requested_cols}")
    missing = [c for c in requested_cols if c not in df.columns]
    if missing:
        print(f"⚠️ Missing columns (will be skipped): {missing}")
    present_cols = [c for c in requested_cols if c in df.columns]
    df = df[present_cols]

    if config.get("drop_empty_rows", False):
        before = df.shape[0]
        df = df.dropna(how="all")
        print(f"🧹 Dropped empty rows: {before - df.shape[0]} removed")

    print_header("STEP 3: ADDING NEW COLUMNS")
    for col, value in config["columns_to_add"].items():
        df[col] = value
        print(f"➕ Added column '{col}' with value: '{value}'")

    print_header("STEP 4: WRITING OUTPUT FILE")
    try:
        df.to_excel(config["output_file"], index=False)
        print(f"✅ Output written to: {config['output_file']} with {df.shape[0]} rows and {df.shape[1]} columns.")
    except Exception as e:
        print(f"❌ Failed to write output file: {e}")
        return

    end_time = time.time()
    print_header("✅ PROCESS COMPLETED")
    print(f"🕒 Finished at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"⏱️ Execution Time: {end_time - start_time:.2f} seconds\n")

# === Run Script ===
if __name__ == "__main__":
    process_excel(CONFIG)
