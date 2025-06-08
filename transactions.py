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
        "App ID", "Application Name", "Scan Date", "Status", "Severity", "Age (days)", "Lifecycle"
    ],

    "columns_to_add": {
        "Reviewed By": "Security Team",
        "Lifecycle": ""  # Used for both EOL and SLO output
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
        {"severity": "Critical", "age_gt": 30, "true_value": "Out of SLO", "false_value": "Ignore"},
        {"severity": "High",     "age_gt": 60, "true_value": "Out of SLO", "false_value": "Ignore"},
        {"severity": "Medium",   "age_gt": 90, "true_value": "Out of SLO", "false_value": "Ignore"},
        {"severity": "Low",      "age_gt": 0,  "true_value": "Ignore",     "false_value": "Ignore"}
    ],

    "drop_empty_rows": True
}

# === FUNCTION DEFINITIONS ===
def print_header(title):
    print(f"\n{'='*60}\n🔷 {title}\n{'='*60}")

def process_excel(config):
    start_time = time.time()

    # STEP 1: Load File
    print_header("STEP 1: LOAD EXCEL FILE")
    if not os.path.exists(config["input_file"]):
        print(f"❌ File not found: {config['input_file']}")
        return
    df = pd.read_excel(config["input_file"], sheet_name=config.get("input_sheet", 0))
    print(f"✅ Loaded {df.shape[0]} rows.")

    # STEP 2: Select Required Columns
    print_header("STEP 2: SELECT COLUMNS")
    df = df[[col for col in config["columns_to_extract"] if col in df.columns]]
    if config.get("drop_empty_rows", False):
        before = df.shape[0]
        df = df.dropna(how="all")
        print(f"🧹 Dropped {before - df.shape[0]} empty rows.")

    # STEP 3: Add Columns
    print_header("STEP 3: ADD STATIC/BLANK COLUMNS")
    for col, val in config["columns_to_add"].items():
        df[col] = val
        print(f"➕ Added column '{col}' with default value: '{val}'")

    # STEP 4: Calculate Age (days)
    print_header("STEP 4: CALCULATE AGE")
    base_col = config["derived_column"]["based_on_column"]
    age_col = config["derived_column"]["column_name"]
    today = datetime.today().date()

    df[base_col] = pd.to_datetime(df[base_col], errors='coerce').dt.date
    df[age_col] = df[base_col].apply(lambda d: (today - d).days if pd.notnull(d) else "")
    print(f"🧮 Calculated '{age_col}' from '{base_col}'")

    # STEP 5: Apply Lifecycle (EOL + SLO)
    print_header("STEP 5: APPLY LIFECYCLE & SLO RULES")
    lconf = config["lifecycle_rule"]
    slo_rules = config["slo_rules"]

    def evaluate_lifecycle(row):
        status = str(row.get(lconf["trigger_column"], "")).lower()
        if lconf["keyword"].lower() in status:
            return lconf["value"]

        severity = str(row.get("Severity", "")).strip().lower()
        age = row.get(age_col)
        try:
            age = int(age)
        except:
            return ""

        for rule in slo_rules:
            if rule["severity"].lower() == severity:
                if severity == "low":
                    return rule["true_value"]
                return rule["true_value"] if age > rule["age_gt"] else rule["false_value"]
        return ""

    df[lconf["target_column"]] = df.apply(evaluate_lifecycle, axis=1)
    print(f"✅ 'Lifecycle' column updated using EOL and SLO rules.")

    # STEP 6: Save to Output Excel
    print_header("STEP 6: SAVE OUTPUT")
    df.to_excel(config["output_file"], index=False)
    print(f"✅ Output written to: {config['output_file']}")

    print_header("✅ PROCESS COMPLETE")
    print(f"⏱️ Execution Time: {time.time() - start_time:.2f} seconds")

# === RUN SCRIPT ===
if __name__ == "__main__":
    process_excel(CONFIG)
