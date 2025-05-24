import pandas as pd
import os
from datetime import datetime

# === CONFIGURATION SECTION ===
CONFIG = {
    # Main file containing original data
    "main_file": "main.xlsx",

    # Output file for final results
    "output_file": "updated_main.xlsx",

    # New columns to add at the beginning
    "new_columns": ["Scan Date", "Reviewer", "Remarks"],

    # Optional static values for new columns
    "column_static_values": {
        "Scan Date": "2025-05-24",
        "Reviewer": "Security Team"
    },

    # Lookup matching configuration
    "main_column_to_match": "App ID",              # Column in main file to match
    "main_column_to_fill": "Owner",                # Column in main file to fill with matched values

    "lookup_file": "reference.xlsx",               # Reference file
    "lookup_sheet_name": "Sheet1",                 # Sheet inside reference file
    "lookup_key_column": "Application ID",         # Column to match against in lookup
    "lookup_value_column": "App Owner",            # Column to fetch value from in lookup

    "unmatched_output_file": "unmatched_rows.xlsx",  # File to write unmatched rows
    "unmatched_placeholder": "ID not found"        # Value to fill for unmatched rows
}

# Helper to nicely format duration
def format_duration(duration):
    seconds = duration.total_seconds()
    if seconds < 60:
        return f"{seconds:.2f} seconds"
    elif seconds < 3600:
        return f"{seconds // 60:.0f} minutes {seconds % 60:.0f} seconds"
    else:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = int(seconds % 60)
        return f"{hours} hours {minutes} minutes {seconds} seconds"

def main():
    start_time = datetime.now()
    print("\n🚀 Starting Excel processing...")

    # Step 1: Check that both required files exist
    for file in [CONFIG["main_file"], CONFIG["lookup_file"]]:
        if not os.path.exists(file):
            print(f"❌ File not found: {file}")
            return
    print("✅ All required files found.")

    # Step 2: Load the main Excel file
    print(f"📖 Reading main file: {CONFIG['main_file']}")
    df_main = pd.read_excel(CONFIG["main_file"], engine="openpyxl")

    # Step 3: Add new columns with static values (if configured)
    print(f"➕ Adding new columns at the beginning: {CONFIG['new_columns']}")
    df_new = pd.DataFrame()
    for col in CONFIG["new_columns"]:
        if col in CONFIG["column_static_values"]:
            val = CONFIG["column_static_values"][col]
            df_new[col] = [val] * len(df_main)
            print(f"🧷 Column '{col}' filled with static value: '{val}'")
        else:
            df_new[col] = [""] * len(df_main)
            print(f"⬜ Column '{col}' left blank")

    # Concatenate new columns to the front of the main data
    df_main = pd.concat([df_new, df_main], axis=1)

    # Step 4: Ensure the target column to fill exists
    if CONFIG["main_column_to_fill"] not in df_main.columns:
        print(f"🆕 Column '{CONFIG['main_column_to_fill']}' not found. Creating it.")
        df_main[CONFIG["main_column_to_fill"]] = ""

    # Step 5: Load the lookup file from specified sheet
    print(f"📖 Reading lookup file: {CONFIG['lookup_file']} (Sheet: {CONFIG['lookup_sheet_name']})")
    df_lookup = pd.read_excel(CONFIG["lookup_file"], sheet_name=CONFIG["lookup_sheet_name"], engine="openpyxl")

    # Step 6: Create a lookup dictionary for fast matching
    print("🔧 Creating lookup dictionary...")
    lookup_dict = pd.Series(
        df_lookup[CONFIG["lookup_value_column"]].values,
        index=df_lookup[CONFIG["lookup_key_column"]]
    ).to_dict()

    # Step 7: Fill values using the lookup dictionary
    print(f"🧩 Matching '{CONFIG['main_column_to_match']}' and filling '{CONFIG['main_column_to_fill']}'...")
    unmatched_rows = []
    matched_rows = 0

    for idx, value in df_main[CONFIG["main_column_to_match"]].items():
        if value in lookup_dict:
            df_main.at[idx, CONFIG["main_column_to_fill"]] = lookup_dict[value]
            matched_rows += 1
        else:
            df_main.at[idx, CONFIG["main_column_to_fill"]] = CONFIG["unmatched_placeholder"]
            unmatched_rows.append(idx)

    print(f"✅ Filled {matched_rows} rows from lookup.")
    print(f"⚠️ {len(unmatched_rows)} unmatched rows set to '{CONFIG['unmatched_placeholder']}'")

    # Step 8: Save unmatched rows to a separate Excel file
    if unmatched_rows:
        df_unmatched = df_main.loc[unmatched_rows]
        df_unmatched.to_excel(CONFIG["unmatched_output_file"], index=False)
        print(f"📝 Unmatched rows written to: {CONFIG['unmatched_output_file']}")
    else:
        print("✅ All rows matched. No unmatched rows found.")

    # Step 9: Save the updated main file
    df_main.to_excel(CONFIG["output_file"], index=False)
    print(f"💾 Final output saved to: {CONFIG['output_file']}")

    # Step 10: Print execution time
    duration = datetime.now() - start_time
    print(f"🕒 Execution time: {format_duration(duration)}")
    print("🎉 Script completed successfully!\n")

# Run the main function
if __name__ == "__main__":
    main()
