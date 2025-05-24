import pandas as pd
import os
import warnings
from datetime import datetime, date

# ✅ Suppress openpyxl warnings about Excel validation
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# === CONFIGURATION ===
CONFIG = {
    "main_file": "main.xlsx",
    "output_file": "updated_main.xlsx",
    "summary_sheet_name": "Summary",
    "unmatched_output_file": "unmatched_rows.xlsx",
    "unmatched_placeholder": "ID not found",

    # Static columns to add at the beginning
    "new_columns": ["Scan Date", "Reviewer"],
    "column_static_values": {
        "Scan Date": "2025-05-24",
        "Reviewer": "Security Team"
    },

    # Lookup columns: match from external file and populate target column
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

    # Fill rule: if string is present in one column, fill another column with a fixed value
    "string_fill_rule": {
        "search_column": "Policy Status",
        "target_column": "Lifecycle",
        "search_string": "Outdated",
        "fill_value": "EOL"
    },

    # Date difference rule: calculate (today - date_column) in days
    "date_diff_rule": {
        "date_column": "Last Reviewed Date",
        "target_column": "Days Since Review"
    }
}

# Format execution duration
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

    # 1️⃣ Load main Excel file
    if not os.path.exists(CONFIG["main_file"]):
        print(f"❌ Main file not found: {CONFIG['main_file']}")
        return
    df_main = pd.read_excel(CONFIG["main_file"], engine="openpyxl")

    # 2️⃣ Add static columns at beginning
    print("➕ Adding static columns at beginning:")
    df_static = pd.DataFrame()
    for col in CONFIG["new_columns"]:
        value = CONFIG["column_static_values"].get(col, "")
        df_static[col] = [value] * len(df_main)
        print(f"  - {col}: '{value}'")
    df_main = pd.concat([df_static, df_main], axis=1)

    all_unmatched_indices = set()
    summary_data = []

    # 3️⃣ Perform lookups
    for lookup in CONFIG["lookups"]:
        print(f"\n🔍 Lookup for column: {lookup['target_column']}")

        if lookup["target_column"] not in df_main.columns:
            df_main[lookup["target_column"]] = ""

        if not os.path.exists(lookup["lookup_file"]):
            print(f"❌ Lookup file missing: {lookup['lookup_file']}")
            continue

        df_lookup = pd.read_excel(lookup["lookup_file"], sheet_name=lookup["sheet_name"], engine="openpyxl")
        lookup_dict = pd.Series(
            df_lookup[lookup["lookup_value_column"]].values,
            index=df_lookup[lookup["lookup_key_column"]]
        ).to_dict()

        matched = unmatched = 0
        for idx, val in df_main[lookup["match_column"]].items():
            if val in lookup_dict:
                df_main.at[idx, lookup["target_column"]] = lookup_dict[val]
                matched += 1
            else:
                df_main.at[idx, lookup["target_column"]] = CONFIG["unmatched_placeholder"]
                all_unmatched_indices.add(idx)
                unmatched += 1

        print(f"✅ Matched: {matched}, ❌ Unmatched: {unmatched}")
        summary_data.append({
            "Target Column": lookup["target_column"],
            "Matched Rows": matched,
            "Unmatched Rows": unmatched
        })

    # 4️⃣ Apply string fill rule (e.g., mark 'Outdated' → 'EOL')
    if "string_fill_rule" in CONFIG:
        rule = CONFIG["string_fill_rule"]
        print(f"\n🔎 String rule: if '{rule['search_string']}' in '{rule['search_column']}', fill '{rule['target_column']}' with '{rule['fill_value']}'")
        if rule["target_column"] not in df_main.columns:
            df_main[rule["target_column"]] = ""
        matched_rows = df_main[rule["search_column"]].astype(str).str.contains(rule["search_string"], case=False, na=False)
        df_main.loc[matched_rows, rule["target_column"]] = rule["fill_value"]
        print(f"✅ Filled '{rule['target_column']}' for {matched_rows.sum()} rows")

    # 5️⃣ Apply date difference rule (today - date_column in days)
    if "date_diff_rule" in CONFIG:
        rule = CONFIG["date_diff_rule"]
        print(f"\n📅 Calculating days since '{rule['date_column']}' into '{rule['target_column']}'")
        if rule["target_column"] not in df_main.columns:
            df_main[rule["target_column"]] = ""
        today = date.today()
        df_main[rule["target_column"]] = pd.to_datetime(df_main[rule["date_column"]], errors='coerce').apply(
            lambda d: (today - d.date()).days if pd.notnull(d) else ""
        )
        print("✅ Date difference calculation completed.")

    # 6️⃣ Save unmatched rows
    if all_unmatched_indices:
        df_main.loc[list(all_unmatched_indices)].to_excel(CONFIG["unmatched_output_file"], index=False)
        print(f"\n⚠️ Unmatched rows written to: {CONFIG['unmatched_output_file']}")
    else:
        print("\n✅ All rows matched.")

    # 7️⃣ Save main output file
    df_main.to_excel(CONFIG["output_file"], index=False)
    print(f"💾 Output written to: {CONFIG['output_file']}")

    # 8️⃣ Save summary sheet
    with pd.ExcelWriter(CONFIG["output_file"], engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        pd.DataFrame(summary_data).to_excel(writer, sheet_name=CONFIG["summary_sheet_name"], index=False)
        print(f"📊 Summary sheet added: {CONFIG['summary_sheet_name']}")

    # 9️⃣ Execution time
    print(f"\n🕒 Execution time: {format_duration(datetime.now() - start_time)}")
    print("🎉 Script completed successfully!\n")

if __name__ == "__main__":
    main()
