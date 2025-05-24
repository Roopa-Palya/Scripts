import pandas as pd
import os
import warnings
from datetime import datetime
from openpyxl import load_workbook

# ✅ Suppress openpyxl warning about data validation extension
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# === CONFIGURATION ===
CONFIG = {
    "main_file": "main.xlsx",
    "output_file": "updated_main.xlsx",
    "summary_sheet_name": "Summary",
    "unmatched_output_file": "unmatched_rows.xlsx",
    "unmatched_placeholder": "ID not found",

    # Static columns to add at beginning
    "new_columns": ["Scan Date", "Reviewer"],
    "column_static_values": {
        "Scan Date": "2025-05-24",
        "Reviewer": "Security Team"
    },

    # Lookup configurations (can add more!)
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
        }
    ],

    # String match rule: search "Outdated" in 'Policy Status' and set 'Lifecycle' to 'EOL'
    "string_fill_rule": {
        "search_column": "Policy Status",
        "target_column": "Lifecycle",
        "search_string": "Outdated",
        "fill_value": "EOL"
    }
}

# Format the duration nicely
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

    # 1️⃣ Load the main Excel file
    if not os.path.exists(CONFIG["main_file"]):
        print(f"❌ Main file not found: {CONFIG['main_file']}")
        return
    df_main = pd.read_excel(CONFIG["main_file"], engine="openpyxl")

    # 2️⃣ Add static columns at the beginning
    df_static = pd.DataFrame()
    print("➕ Adding static columns at beginning:")
    for col in CONFIG["new_columns"]:
        val = CONFIG["column_static_values"].get(col, "")
        df_static[col] = [val] * len(df_main)
        print(f"  - {col}: '{val}'")
    df_main = pd.concat([df_static, df_main], axis=1)

    all_unmatched_indices = set()
    summary_data = []

    # 3️⃣ Loop through lookup configs
    for lookup in CONFIG["lookups"]:
        print(f"\n🔍 Processing lookup for column: {lookup['target_column']}")

        # Ensure the target column exists
        if lookup["target_column"] not in df_main.columns:
            df_main[lookup["target_column"]] = ""

        # Load the lookup Excel file
        if not os.path.exists(lookup["lookup_file"]):
            print(f"❌ Lookup file not found: {lookup['lookup_file']}")
            continue
        df_lookup = pd.read_excel(lookup["lookup_file"], sheet_name=lookup["sheet_name"], engine="openpyxl")

        # Create a dictionary to map values
        lookup_dict = pd.Series(
            df_lookup[lookup["lookup_value_column"]].values,
            index=df_lookup[lookup["lookup_key_column"]]
        ).to_dict()

        matched, unmatched = 0, 0

        for idx, value in df_main[lookup["match_column"]].items():
            if value in lookup_dict:
                df_main.at[idx, lookup["target_column"]] = lookup_dict[value]
                matched += 1
            else:
                df_main.at[idx, lookup["target_column"]] = CONFIG["unmatched_placeholder"]
                all_unmatched_indices.add(idx)
                unmatched += 1

        print(f"✅ Matched: {matched}, ❌ Unmatched: {unmatched}")

        summary_data.append({
            "Target Column": lookup["target_column"],
            "Match Column": lookup["match_column"],
            "Lookup File": lookup["lookup_file"],
            "Sheet": lookup["sheet_name"],
            "Matched Rows": matched,
            "Unmatched Rows": unmatched
        })

    # 4️⃣ Apply string match fill rule
    if "string_fill_rule" in CONFIG:
        rule = CONFIG["string_fill_rule"]
        print(f"\n🔎 Applying string match rule: If '{rule['search_string']}' in '{rule['search_column']}', then set '{rule['target_column']}' to '{rule['fill_value']}'")
        if rule["target_column"] not in df_main.columns:
            df_main[rule["target_column"]] = ""
        matched_rows = df_main[rule["search_column"]].astype(str).str.contains(rule["search_string"], case=False, na=False)
        df_main.loc[matched_rows, rule["target_column"]] = rule["fill_value"]
        print(f"✅ Marked {matched_rows.sum()} rows as '{rule['fill_value']}' in '{rule['target_column']}'")

    # 5️⃣ Save unmatched rows to file
    if all_unmatched_indices:
        df_main.loc[list(all_unmatched_indices)].to_excel(CONFIG["unmatched_output_file"], index=False)
        print(f"\n⚠️ Unmatched rows saved to: {CONFIG['unmatched_output_file']}")
    else:
        print("\n✅ No unmatched rows found.")

    # 6️⃣ Save final output Excel
    df_main.to_excel(CONFIG["output_file"], index=False)
    print(f"💾 Final output saved to: {CONFIG['output_file']}")

    # 7️⃣ Append summary sheet
    with pd.ExcelWriter(CONFIG["output_file"], engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        pd.DataFrame(summary_data).to_excel(writer, sheet_name=CONFIG["summary_sheet_name"], index=False)
        print(f"📊 Lookup summary written to sheet: {CONFIG['summary_sheet_name']}")

    # 8️⃣ Execution time
    print(f"\n🕒 Execution time: {format_duration(datetime.now() - start_time)}")
    print("🎉 Script completed successfully!\n")

if __name__ == "__main__":
    main()
