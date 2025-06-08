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
        "App ID", "Application Name", "Scan Date", "Status", "Severity", "Lifecycle", "Flag"
    ],

    "columns_to_add": {
        "Reviewed By": "Security Team",
        "Lifecycle": "",
        "Flag": ""
    },

    "inplace_age_check": {
        "date_column": "Scan Date",
        "age_variable": "age_days"
    },

    "lifecycle_rules": {
        "target_column": "Lifecycle",
        "eol_rule": {
            "trigger_column": "Status",
            "keyword": "Outdated",
            "value": "EOL"
        },
        "slo_rules": [
            {"severity": "Critical",      "age_gt": 30,  "true_value": "Out of SLO", "false_value": "Ignore"},
            {"severity": "High",          "age_gt": 60,  "true_value": "Out of SLO", "false_value": "Ignore"},
            {"severity": "Medium",        "age_gt": 90,  "true_value": "Out of SLO", "false_value": "Ignore"},
            {"severity": "Low",           "age_gt": 0,   "true_value": "Ignore",     "false_value": "Ignore"},
            {"severity": "Informational", "age_gt": 0,   "true_value": "Ignore",     "false_value": "Ignore"}
        ]
    },

    "final_mapping": {
        "source_column": "Lifecycle",
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
    print(f"\n{'=' * 60}\n🔷 {title}\n{'=' * 60}")

def process_excel(config):
    start_time = time.time()

    # STEP 1: Load input Excel
    print_header("STEP 1: LOAD FILE")
    if not os.path.exists(config["input_file"]):
        print(f"❌ File not found: {config['input_file']}")
        return
    df = pd.read_excel(config["input_file"], sheet_name=config["input_sheet"])
    print(f"✅ Loaded {df.shape[0]} rows.")

    # STEP 2: Keep required columns
    print_header("STEP 2: FILTER COLUMNS")
    df = df[[col for col in config["columns_to_extract"] if col in df.columns]]
    if config.get("drop_empty_rows"):
        before = df.shape[0]
        df = df.dropna(how="all")
        print(f"🧹 Dropped {before - df.shape[0]} empty rows.")

    # STEP 3: Add static/blank columns
    print_header("STEP 3: ADD COLUMNS")
    for col, val in config["columns_to_add"].items():
        df[col] = val
        print(f"➕ Added column: {col}")

    # STEP 4: Calculate age (internally)
    print_header("STEP 4: CALCULATE AGE")
    age_cfg = config["inplace_age_check"]
    today = datetime.today().date()
    df[age_cfg["date_column"]] = pd.to_datetime(df[age_cfg["date_column"]], errors='coerce').dt.date
    df[age_cfg["age_variable"]] = df[age_cfg["date_column"]].apply(
        lambda d: (today - d).days if pd.notnull(d) else -1
    )
    print("🧮 Age calculated internally.")

    # STEP 5: Lifecycle classification
    print_header("STEP 5: APPLY LIFECYCLE LOGIC")
    rules = config["lifecycle_rules"]
    tgt = rules["target_column"]
    eol = rules["eol_rule"]
    slo = rules["slo_rules"]

    def classify(row):
        if eol["keyword"].lower() in str(row.get(eol["trigger_column"], "")).lower():
            return eol["value"]
        severity = str(row.get("Severity", "")).strip().lower()
        age = row.get(age_cfg["age_variable"])
        try:
            age = int(age)
        except:
            return ""
        for rule in slo:
            if rule["severity"].lower() == severity:
                return rule["true_value"] if age > rule["age_gt"] else rule["false_value"]
        return ""

    df[tgt] = df.apply(classify, axis=1)
    print(f"✅ '{tgt}' filled using severity + age logic.")

    # STEP 6: Final mapping into separate target column
    print_header("STEP 6: FINAL MAPPING TO TARGET COLUMN")
    fmap = config["final_mapping"]
    src = fmap["source_column"]
    dst = fmap["target_column"]
    df[dst] = df[src].map(fmap["map_values"]).fillna("")
    print(f"🔁 Final values written to '{dst}' from '{src}'")

    # STEP 7: Save
    print_header("STEP 7: SAVE TO FILE")
    df.drop(columns=[age_cfg["age_variable"]], errors='ignore').to_excel(config["output_file"], index=False)
    print(f"✅ Output saved: {config['output_file']}")

    print_header("✅ DONE")
    print(f"⏱️ Completed in {time.time() - start_time:.2f} seconds")

# === RUN ===
if __name__ == "__main__":
    process_excel(CONFIG)
