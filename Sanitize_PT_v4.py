import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook

# === CONFIGURATION ===
CONFIG = {
    "main_file": "main.xlsx",
    "output_file": "updated_main.xlsx",
    "summary_sheet_name": "Summary",
    "unmatched_output_file": "unmatched_rows.xlsx",
    "unmatched_placeholder": "ID not found",

    "new_columns": ["Scan Date", "Reviewer"],
    "column_static_values": {
        "Scan Date": "2025-05-24",
        "Reviewer": "Security Team"
    },

    # 🔁 List of all lookup mappings
    "lookups": [
        {
            "target_column": "Owner",
            "match_column": "App ID",
            "lookup_file": "owners.xlsx",
            "sheet_name": "Sheet1",
            "lookup_key_column": "Application ID",
            "lookup_value_column": "App Owner"
        },
        {
            "target_column": "Business Unit",
            "match_column": "App ID",
            "lookup_file": "bu.xlsx",
            "sheet_name": "Sheet1",
            "lookup_key_column": "App Identifier",
            "lookup_value_column": "BU Name"
        },
        {
            "target_column": "Location",
            "match_column": "App ID",
            "lookup_file": "locations.xlsx",
            "sheet_name": "Sheet1",
            "lookup_key_column": "App ID",
            "lookup_value_column": "Region"
        },
        {
            "target_column": "App Type",
            "match_column": "App ID",
            "lookup_file": "types.xlsx",
            "sheet_name": "Sheet1",
            "lookup_key_column": "ID",
            "lookup_value_column": "Type"
        },
        {
            "target_column": "Risk Rating",
            "match_column": "App ID",
            "lookup_file": "risk_data.xlsx",
            "sheet_name": "Sheet1",
            "lookup_key_column": "App ID",
            "lookup_value_column": "Risk"
        }
    ]
}

# 🕒 Format elapsed time nicely
def format_duration(duration):
    seconds = duration.total_seconds()
    if seconds < 60:
        return f"{seconds:.2f} seconds"
    elif seconds < 3600:
        return f"{seconds // 60:.0f} minutes {seconds % 60:.0f} seconds"
    else:
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        s = int(seconds % 60)
        return f"{h}h {m}m {s}s"

def main():
    start_time = datetime.now()
    print("\n🚀 Starting Excel enhancement script...\n")

    # Step 1: Read the main Excel file
    if not os.path.exists(CONFIG["main_file"]):
        print(f"❌ Main file not found: {CONFIG['main_file']}")
        return
    df_main = pd.read_excel(CONFIG["main_file"], engine="openpyxl")

    # Step 2: Add static columns at the beginning
    df_static = pd.DataFrame()
    print("➕ Adding static columns at beginning:")
    for col in CONFIG["new_columns"]:
        val = CONFIG["column_static_values"].get(col, "")
        df_static[col] = [val] * len(df_main)
        print(f"  - {col}: '{val}'")
    df_main = pd.concat([df_static, df_main], axis=1)

    all_unmatched_indices = set()
    summary_data = []  # to hold matched/unmatched info for summary tab

    # Step 3: Process each lookup entry
    for lookup in CONFIG["lookups"]:
        print(f"\n🔍 Processing lookup for: {lookup['target_column']}")

        # Ensure target column exists
        if lookup["target_column"] not in df_main.columns:
            df_main[lookup["target_column"]] = ""

        # Load lookup Excel
        if not os.path.exists(lookup["lookup_file"]):
            print(f"❌ Lookup file not found: {lookup['lookup_file']}")
            continue

        df_lookup = pd.read_excel(lookup["lookup_file"], sheet_name=lookup["sheet_name"], engine="openpyxl")

        # Build lookup dictionary
        lookup_dict = pd.Series(
            df_lookup[lookup["lookup_value_column"]].values,
            index=df_lookup[lookup["lookup_key_column"]]
        ).to_dict()

        matched, unmatched = 0, 0

        # Fill values in main file
        for idx, value in df_main[lookup["match_column"]].items():
            if value in lookup_dict:
                df_main.at[idx, lookup["target_column"]] = lookup_dict[value]
                matched += 1
            else:
                df_main.at[idx, lookup["target_column"]] = CONFIG["unmatched_placeholder"]
                all_unmatched_indices.add(idx)
                unmatched += 1

        print(f"✅ Matched: {matched}, ❌ Unmatched: {unmatched}")

        # Add to summary
        summary_data.append({
            "Target Column": lookup["target_column"],
            "Match Column": lookup["match_column"],
            "Lookup File": lookup["lookup_file"],
            "Sheet": lookup["sheet_name"],
            "Matched Rows": matched,
            "Unmatched Rows": unmatched
        })

    # Step 4: Save unmatched rows if any
    if all_unmatched_indices:
        df_main.loc[list(all_unmatched_indices)].to_excel(CONFIG["unmatched_output_file"], index=False)
        print(f"\n⚠️ Unmatched rows saved to: {CONFIG['unmatched_output_file']}")
    else:
        print("\n✅ No unmatched rows found.")

    # Step 5: Save main file
    df_main.to_excel(CONFIG["output_file"], index=False)
    print(f"💾 Final enhanced file saved to: {CONFIG['output_file']}")

    # Step 6: Append Summary Sheet
    with pd.ExcelWriter(CONFIG["output_file"], engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        pd.DataFrame(summary_data).to_excel(writer, sheet_name=CONFIG["summary_sheet_name"], index=False)
        print(f"📊 Lookup summary saved to sheet: {CONFIG['summary_sheet_name']}")

    print(f"\n🕒 Execution time: {format_duration(datetime.now() - start_time)}")
    print("🎉 Script completed successfully!\n")

# Entry point
if __name__ == "__main__":
    main()
