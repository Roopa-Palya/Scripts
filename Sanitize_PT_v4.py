import pandas as pd
import os
import warnings
from datetime import datetime, date

# ✅ Suppress Excel warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# === CONFIGURATION ===
CONFIG = {
    "main_file": "main.xlsx",
    "output_file": "updated_main.xlsx",
    "summary_sheet_name": "Summary",
    "unmatched_output_file": "unmatched_rows.xlsx",
    "unmatched_placeholder": "ID not found",

    # Static columns to prepend
    "new_columns": ["Scan Date", "Reviewer"],
    "column_static_values": {
        "Scan Date": "2025-05-24",
        "Reviewer": "Security Team"
    },

    # Lookup from external files
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

    # String match: fill if keyword found
    "string_fill_rule": {
        "search_column": "Policy Status",
        "target_column": "Lifecycle",
        "search_string": "Outdated",
        "fill_value": "EOL"
    },

    # Days since date column
    "date_diff_rule": {
        "date_column": "Last Reviewed Date",
        "target_column": "Days Since Review"
    },

    # Severity vs Count rule
    "severity_count_rule": {
        "severity_column": "Severity",
        "count_column": "Count",
        "target_column": "Status",
        "rules": {
            "Critical": 30,
            "High": 60,
            "Medium": 90,
            "Low": 180
        },
        "value_if_true": "OOS",
        "value_if_false": "WIS"
    }
}

# Format execution time
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

# Main script logic
def main():
    start_time = datetime.now()
    print("\n🚀 Starting Excel enhancement script...\n")

    # Load main Excel file
    if not os.path.exists(CONFIG["main_file"]):
        print(f"❌ File not found: {CONFIG['main_file']}")
        return
    df_main = pd.read_excel(CONFIG["main_file"], engine="openpyxl")

    # Add static columns
    print("➕ Adding static columns:")
    df_static = pd.DataFrame()
    for col in CONFIG["new_columns"]:
        val = CONFIG["column_static_values"].get(col, "")
        df_static[col] = [val] * len(df_main)
        print(f"  - {col}: '{val}'")
    df_main = pd.concat([df_static, df_main], axis=1)

    all_unmatched_indices = set()
    summary_data = []

    # Apply lookups
    for lookup in CONFIG["lookups"]:
        print(f"\n🔍 Lookup for: {lookup['target_column']}")
        if lookup["target_column"] not in df_main.columns:
            df_main[lookup["target_column"]] = ""

        if not os.path.exists(lookup["lookup_file"]):
            print(f"❌ Missing file: {lookup['lookup_file']}")
            continue

        df_lookup = pd.read_excel(lookup["lookup_file"], sheet_name=lookup["sheet_name"], engine="openpyxl")
        lookup_dict = pd.Series(
            df_lookup[lookup["lookup_value_column"]].values,
            index=df_lookup[lookup["lookup_key_column"]]
        ).to_dict()

        matched = unmatched = 0
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
            "Matched Rows": matched,
            "Unmatched Rows": unmatched
        })

    # String fill rule
    if "string_fill_rule" in CONFIG:
        rule = CONFIG["string_fill_rule"]
        print(f"\n🔎 Filling '{rule['target_column']}' if '{rule['search_string']}' in '{rule['search_column']}'")
        if rule["target_column"] not in df_main.columns:
            df_main[rule["target_column"]] = ""
        matched_rows = df_main[rule["search_column"]].astype(str).str.contains(rule["search_string"], case=False, na=False)
        df_main.loc[matched_rows, rule["target_column"]] = rule["fill_value"]
        print(f"✅ Filled {matched_rows.sum()} rows")

    # Date difference rule
    if "date_diff_rule" in CONFIG:
        rule = CONFIG["date_diff_rule"]
        print(f"\n📅 Calculating days since '{rule['date_column']}' → '{rule['target_column']}'")
        if rule["target_column"] not in df_main.columns:
            df_main[rule["target_column"]] = ""
        today = date.today()
        df_main[rule["target_column"]] = pd.to_datetime(df_main[rule["date_column"]], errors="coerce").apply(
            lambda d: (today - d.date()).days if pd.notnull(d) else ""
        )
        print("✅ Date difference calculated")

    # Severity + Count rule
    if "severity_count_rule" in CONFIG:
        rule = CONFIG["severity_count_rule"]
        print(f"\n📊 Applying severity-count logic → '{rule['target_column']}'")
        if rule["target_column"] not in df_main.columns:
            df_main[rule["target_column"]] = ""

        for severity, threshold in rule["rules"].items():
            condition = (
                (df_main[rule["severity_column"]] == severity) &
                (df_main[rule["count_column"]] > threshold)
            )
            df_main.loc[condition, rule["target_column"]] = rule["value_if_true"]
            df_main.loc[(df_main[rule["severity_column"]] == severity) & ~condition, rule["target_column"]] = rule["value_if_false"]
        print("✅ Status filled based on Severity and Count")

    # Save unmatched rows
    if all_unmatched_indices:
        df_main.loc[list(all_unmatched_indices)].to_excel(CONFIG["unmatched_output_file"], index=False)
        print(f"\n⚠️ Saved unmatched rows to: {CONFIG['unmatched_output_file']}")
    else:
        print("\n✅ All lookups matched")

    # Save final Excel
    df_main.to_excel(CONFIG["output_file"], index=False)
    print(f"💾 Final file saved: {CONFIG['output_file']}")

    # Save summary sheet
    with pd.ExcelWriter(CONFIG["output_file"], engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        pd.DataFrame(summary_data).to_excel(writer, sheet_name=CONFIG["summary_sheet_name"], index=False)
        print(f"📊 Summary added: {CONFIG['summary_sheet_name']}")

    # Final execution log
    print(f"\n🕒 Execution time: {format_duration(datetime.now() - start_time)}")
    print("🎉 Script finished successfully!\n")

if __name__ == "__main__":
    main()
