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
        "App ID", "Application Name", "Scan Date", "Status", "Severity", "Age (days)", "SLO Status", "Lifecycle"
    ],

    "columns_to_add": {
        "Reviewed By": "Security Team",
        "SLO Status": "",      # Will be filled
        "Lifecycle": ""        # Will be filled
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

# === UTILITY ===
def print_header(title):
    print(f"\n{'=' * 60}\n🔷 {title}\n{'=' * 60}")

# === MAIN PROCESS ===
def process_excel(config):
    start_time = time.time()

    # STEP 1: Load File
    print_header("STEP 1: LOADING EXCEL FILE")
    if not os.path.exists(config["input_file"]):
        print(f"❌ File not found: {config['input_file']}")
        return
    df = pd.read_excel(config["input_file"], sheet_name=config.get("input_sheet", 0))
    print(f"✅ Loaded {df.shape[0]} rows.")

    # STEP 2: Extract Columns
    print_header("STEP 2: SELECTING REQUIRED COLUMNS")
    df = df[[col for col in config["columns_to_extract"] if col in df.columns]]
    if config.get("drop_empty_rows", False):
        before = df.shape[0]
        df = df.dropna(how="all")
        print(f"🧹 Dropped {before - df.shape[0]} empty rows.")

    # STEP 3: Add New Columns
    print_header("STEP 3: ADDING NEW COLUMNS")
    for col, val in config["columns_to_add"].items():
        df[col] = val
        print(f"➕ Column added: {col} = '{val}'")

    # STEP 4: Calculate Age (days)
    print_header("STEP 4: CALCULATING AGE COLUMN")
    dconf = config["derived_column"]
    date_col = dconf["based_on_column"]
    age_col = dconf["column_name"]
    today = datetime.today().date()

    df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.date
    df[age_col] = df[date_col].apply(lambda d: (today - d).days if pd.notnull(d) else "")
    print(f"🧮 Age calculated in column '{age_col}' based on '{date_col}'")

    # STEP 5: Lifecycle Rule (Outdated = EOL)
    print_header("STEP 5: APPLYING LIFECYCLE RULE")
    lconf = config["lifecycle_rule"]
    df[lconf["target_column"]] = df.apply(
        lambda row: lconf["value"] if lconf["keyword"].lower() in str(row.get(lconf["trigger_column"], "")).lower()
        else row.get(lconf["target_column"], ""),
        axis=1
    )
    print(f"📎 Lifecycle rule applied: if '{lconf['trigger_column']}' contains '{lconf['keyword']}' → '{lconf['target_column']}' = '{lconf['value']}'")

    # STEP 6: Apply SLO Rules
    print_header("STEP 6: APPLYING SLO RULES")
    def evaluate_slo(row):
        severity = str(row.get("Severity", "")).strip().lower()
        age = row.get(age_col)
        try:
            age = int(age)
        except:
            return ""

        for rule in config["slo_rules"]:
            if rule["severity"].lower() == severity:
                if severity == "low":
                    return rule["true_value"]
                return rule["true_value"] if age > rule["age_gt"] else rule["false_value"]
        return ""

    df["SLO Status"] = df.apply(evaluate_slo, axis=1)
    print(f"✅ SLO rules applied for Severity + Age")

    # STEP 7: Save Output
    print_header("STEP 7: SAVING OUTPUT")
    df.to_excel(config["output_file"], index=False)
    print(f"✅ Output written to: {config['output_file']}")

    print_header("✅ DONE")
    print(f"⏱️ Total Execution Time: {time.time() - start_time:.2f} seconds")

# === ENTRY POINT ===
if __name__ == "__main__":
    process_excel(CONFIG)
