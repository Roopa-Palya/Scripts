import pandas as pd
import os
from datetime import datetime

# === CONFIGURATION SECTION ===
CONFIG = {
    "main_file": "main.xlsx",  # Original data file
    "output_file": "updated_main.xlsx",  # Final output file
    "new_columns": ["Scan Date", "Reviewer", "Remarks"],  # Columns to add at the beginning
    "column_static_values": {  # Static values for new columns
        "Scan Date": "2025-05-24",
        "Reviewer": "Security Team"
    },

    # Lookup configuration
    "main_column_to_match": "App ID",              # Column in main file to match
    "main_column_to_fill": "Owner",                # Column in main file to fill
    "lookup_file": "reference.xlsx",               # File to lookup from
    "lookup_sheet_name": "Sheet1",                 # Sheet name in the lookup file
    "lookup_key_column": "Application ID",         # Column in lookup file to match
    "lookup_value_column": "App Owner",            # Value to fill from lookup
    "unmatched_output_file": "unmatched_rows.xlsx" # Output file for unmatched rows
}

# Function to nicely format execution time
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

# === MAIN FUNCTION START ===
def main():
    start_time = datetime.now()
    print("\n🚀 Starting Excel processing...")

    # 1. Check both required files exist
    for file in [CONFIG["main_file"], CONFIG["lookup_file"]]:
        if not os.path.exists(file):
            print(f"❌ File not found: {file}")
            return
    print("✅ All required files found.")

    # 2. Load main Excel file
    print(f"📖 Reading main file: {CONFIG['main_file']}")
    df_main = pd.read_excel(CONFIG["main_file"], engine="openpyxl")

    # 3. Add new columns with static values (or leave blank)
    print(f"➕ Adding new columns at the beginning: {CONFIG['new_columns']}")
    df_new = pd.DataFrame()

    for col in CONFIG["new_columns"]:
        if col in CONFIG["column_static_values"]:
            val = CONFIG["column_static_values"][col]
            df_new[col] = [val] * len(df_main)
            print(f"🧷 Column '{col}' set to static value: {val}")
        else:
            df_new[col] = [""] * len(df_main)
            print(f"⬜ Column '{col}' left blank.")

    # Merge new columns with original data
    df_main = pd.concat([df_new, df_main], axis=1)

    # 4. Add the target column to fill (if not already present)
    if CONFIG["main_column_to_fill"] not in df_main.columns:
        print(f"🆕 Creating column: {CONFIG['main_column_to_fill']}")
        df_main[CONFIG["main_column_to_fill"]] = ""

    # 5. Load lookup Excel (specific sheet)
    print(f"📖 Reading lookup file: {CONFIG['lookup_file']} (Sheet: {CONFIG['lookup_sheet_name']})")
    df_lookup = pd.read_excel(CONFIG["lookup_file"], sheet_name=CONFIG["lookup_sheet_name"], engine="openpyxl")

    # 6. Create a lookup dictionary
    print("🔧 Creating lookup mapping...")
    lookup_dict = pd.Series(
        df_lookup[CONFIG["lookup_value_column"]].values,
        index=df_lookup[CONFIG["lookup_key_column"]]
    ).to_dict()

    # 7. Fill values into main column using the lookup
    print(f"🧩 Matching '{CONFIG['main_column_to_match']}' and filling '{CONFIG['main_column_to_fill']}'")
    unmatched_rows = []
    matched_rows = 0

    for idx, val in df_main[CONFIG["main_column_to_match"]].items():
        if val in lookup_dict:
            df_main.at[idx, CONFIG["main_column_to_fill"]] = lookup_dict[val]
            matched_rows += 1
        else:
            unmatched_rows.append(idx)

    print(f"✅ Filled {matched_rows} rows from reference data.")

    # 8. Save unmatched rows to a separate file
    if unmatched_rows:
        df_unmatched = df_main.loc[unmatched_rows]
        df_unmatched.to_excel(CONFIG["unmatched_output_file"], index=False)
        print(f"⚠️ Unmatched rows saved to: {CONFIG['unmatched_output_file']} ({len(unmatched_rows)} rows)")
    else:
        print("✅ All rows matched. No unmatched rows found.")

    # 9. Save the final updated file
    df_main.to_excel(CONFIG["output_file"], index=False)
    print(f"💾 Final output written to: {CONFIG['output_file']}")

    # 10. Print execution time
    duration = datetime.now() - start_time
    print(f"🕒 Execution time: {format_duration(duration)}")
    print("🎉 Script completed successfully!\n")

# Entry point
if __name__ == "__main__":
    main()
