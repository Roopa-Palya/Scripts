import pandas as pd
import os
import time
from datetime import datetime

# === CONFIGURATION ===
CONFIG = {
    "input_file": "input_data.xlsx",
    "input_sheet": "Sheet1",
    "output_file": "processed_output.xlsx",

    "columns_to_extract": [
        "App ID", "Application Name", "Scan Date", "Status", "Severity", "Age (days)", "Lifecycle", "Flag"
    ],

    "columns_to_add": {
        "Reviewed By": "Security Team",
        "Lifecycle": "",
        "Flag": ""
    },

    "derived_column": {
        "column_name": "Age (days)",
        "based_on_column": "Scan Date"
    },

    "lifecycle_rule": {
        "trigger_column": "Status",
        "keyword": "Outdated",
        "target_column": "Lifecycle",
        "value": "EOL"
    },

    "slo_rules": [
        {"severity": "Critical",       "age_gt": 30,  "true_value": "Out of SLO", "false_value": "Ignore"},
        {"severity": "High",           "age_gt": 60,  "true_value": "Out of SLO", "false_value": "Ignore"},
        {"severity": "Medium",         "age_gt": 90,  "true_value": "Out of SLO", "false_value": "Ignore"},
        {"severity": "Low",            "age_gt": 0,   "true_value": "Ignore",     "false_value": "Ignore"},
        {"severity": "Informational",  "age_gt": 0,   "true_value": "Ignore",     "false_value": "Ignore"}
    ],

    "final_mapping": {
        "column_to_map": "Lifecycle",
        "target_column": "Flag",
        "map_values": {
            "EOL": "EOL",
            "Ignore": 0,
            "Out of SLO": 1
        }
    },

    "drop_empty_rows": True
}

# === MAIN FUNCTION ===
def print_header(title):
    print(f"\n{'='*60}\n🔷 {title}\n{'='*60}")

def process_excel(config):
    start_time = time.time()

    # STEP 1: Load Excel
    print_header("STEP 1: LOAD FILE")
    if not os.path.exists(config["input_file"]):
        print(f"❌ File not found: {config['input_file']}")
        return
    df = pd.read_excel(config["input_file"], sheet_name=config.get("input_sheet", 0))
    print(f"✅ Loaded {df.shape[0]} rows.")

    # STEP 2: Keep only required columns
    print_header("STEP 2: SELECT COLUMNS")
    df = df[[col for col in config["columns_to_extract"] if col in df.columns]]
    if config.get("drop_empty_rows", False):
        before = df.shape[0]
        df = df.dropna(how="all")
        print(f"🧹 Dropped {before - df.shape[0]} empty rows.")

    # STEP 3: Add new columns
    print_header("STEP 3: ADD COLUMNS")
    for col, val in config["columns_to_add"].items():
        df[col] = val
        print(f"➕ Added column: {col} = '{val}'")

    # STEP 4: Calculate Age in Days
    print_header("STEP 4: CALCULATE AGE")
    age_col = config["derived_column"]["column_name"]
    base_col = config["derived_column"]["based_on_column"]
    today = datetime.today().date()
    df[base_col] = pd.to_datetime(df[base_col], errors='coerce').dt.date
    df[age_col] = df[base_col].apply(lambda d: (today - d).days if pd.notnull(d) else "")
    print(f"🧮 Calculated '{age_col}' from '{base_col}'")

    # STEP 5: Apply Lifecycle (EOL and SLO combined)
    print_header("STEP 5: SET LIFECYCLE")
    lconf = config["lifecycle_rule"]
    slo_rules = config["slo_rules"]

    def set_lifecycle(row):
        if lconf["keyword"].lower() in str(row.get(lconf["trigger_column"], "")).lower():
            return lconf["value"]
        severity = str(row.get("Severity", "")).strip().lower()
        age = row.get(age_col)
        try:
            age = int(age)
        except:
            return ""
        for rule in slo_rules:
            if rule["severity"].lower() == severity:
                if age > rule["age_gt"]:
                    return rule["true_value"]
                else:
                    return rule["false_value"]
        return ""

    df[lconf["target_column"]] = df.apply(set_lifecycle, axis=1)
    print(f"✅ Lifecycle values set using SLO and status rules.")

    # STEP 6: Final Mapping (e.g. Lifecycle → Flag)
    print_header("STEP 6: FINAL FLAG MAPPING")
    fmap = config["final_mapping"]
    source_col = fmap["column_to_map"]
    target_col = fmap["target_column"]
    mapping = fmap["map_values"]

    df[target_col] = df[source_col].map(mapping).fillna("")
    print(f"✅ Final column '{target_col}' mapped from '{source_col}'")

    # STEP 7: Save to file
    print_header("STEP 7: SAVE OUTPUT")
    df.to_excel(config["output_file"], index=False)
    print(f"✅ Output saved: {config['output_file']}")

    print_header("✅ COMPLETE")
    print(f"⏱️ Execution Time: {time.time() - start_time:.2f} seconds")

# === RUN SCRIPT ===
if __name__ == "__main__":
    process_excel(CONFIG)
