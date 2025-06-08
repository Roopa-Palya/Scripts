import pandas as pd
import os
import time
from datetime import datetime

# === CONFIGURATION ===
CONFIG = {
    "input_file": "input_data.xlsx",        # Path to your input Excel file
    "input_sheet": "Sheet1",                # Sheet name to read from
    "output_file": "processed_output.xlsx", # Path to your output Excel file

    "columns_to_extract": [                 # Columns to retain from the input
        "App ID", "Application Name", "Scan Date", "Status", "Severity", "Lifecycle"
    ],

    "columns_to_add": {                     # New columns to add (overwrites if exists)
        "Reviewed By": "Security Team",
        "Lifecycle": "",
    },

    "inplace_age_check": {                  # Calculate age (today - date column) — internal only
        "date_column": "Scan Date",
        "age_variable": "age_days"
    },

    "lifecycle_rules": {
        "target_column": "Lifecycle",       # Column to be updated with logic
        "eol_rule": {                       # Rule: if Status contains "Outdated"
            "trigger_column": "Status",
            "keyword": "Outdated",
            "value": "EOL"
        },
        "slo_rules": [                      # Severity + Age rules
            {"severity": "Critical",      "age_gt": 30,  "true_value": "Out of SLO", "false_value": "Ignore"},
            {"severity": "High",          "age_gt": 60,  "true_value": "Out of SLO", "false_value": "Ignore"},
            {"severity": "Medium",        "age_gt": 90,  "true_value": "Out of SLO", "false_value": "Ignore"},
            {"severity": "Low",           "age_gt": 0,   "true_value": "Ignore",     "false_value": "Ignore"},
            {"severity": "Informational", "age_gt": 0,   "true_value": "Ignore",     "false_value": "Ignore"}
        ]
    },

    "final_mapping": {                      # Final transformation of values in Lifecycle
        "column": "Lifecycle",
        "map_values": {
            "EOL": "EOL",
            "Ignore": 0,
            "Out of SLO": 1
        }
    },

    "drop_empty_rows": True
}

# === PROCESSING LOGIC ===
def print_header(title):
    print(f"\n{'=' * 60}\n🔷 {title}\n{'=' * 60}")

def process_excel(config):
    start_time = time.time()

    # STEP 1: Load file
    print_header("STEP 1: LOAD INPUT FILE")
    if not os.path.exists(config["input_file"]):
        print(f"❌ File not found: {config['input_file']}")
        return
    df = pd.read_excel(config["input_file"], sheet_name=config["input_sheet"])
    print(f"✅ Loaded {df.shape[0]} rows.")

    # STEP 2: Filter required columns
    print_header("STEP 2: SELECT COLUMNS")
    df = df[[col for col in config["columns_to_extract"] if col in df.columns]]
    if config.get("drop_empty_rows"):
        before = df.shape[0]
        df = df.dropna(how="all")
        print(f"🧹 Removed {before - df.shape[0]} empty rows.")

    # STEP 3: Add static/blank columns
    print_header("STEP 3: ADD COLUMNS")
    for col, val in config["columns_to_add"].items():
        df[col] = val
        print(f"➕ Column added: {col} = '{val}'")

    # STEP 4: Calculate age difference (internal use only)
    print_header("STEP 4: CALCULATE AGE IN DAYS")
    age_cfg = config["inplace_age_check"]
    today = datetime.today().date()
    df[age_cfg["date_column"]] = pd.to_datetime(df[age_cfg["date_column"]], errors='coerce').dt.date
    df[age_cfg["age_variable"]] = df[age_cfg["date_column"]].apply(
        lambda d: (today - d).days if pd.notnull(d) else -1
    )
    print("🧮 Internal 'age_days' calculated.")

    # STEP 5: Apply lifecycle logic (EOL + Severity rules)
    print_header("STEP 5: APPLY LIFECYCLE RULES")
    rules = config["lifecycle_rules"]
    tgt_col = rules["target_column"]
    eol = rules["eol_rule"]
    slo_rules = rules["slo_rules"]

    def classify(row):
        if eol["keyword"].lower() in str(row.get(eol["trigger_column"], "")).lower():
            return eol["value"]
        severity = str(row.get("Severity", "")).strip().lower()
        age = row.get(age_cfg["age_variable"])
        try:
            age = int(age)
        except:
            return ""
        for rule in slo_rules:
            if rule["severity"].lower() == severity:
                return rule["true_value"] if age > rule["age_gt"] else rule["false_value"]
        return ""

    df[tgt_col] = df.apply(classify, axis=1)
    print(f"✅ '{tgt_col}' populated with lifecycle classification.")

    # STEP 6: Final overwrite mapping on same column
    print_header("STEP 6: FINAL MAPPING IN-PLACE")
    fmap = config["final_mapping"]
    df[fmap["column"]] = df[fmap["column"]].map(fmap["map_values"]).fillna("")
    print(f"🔁 Final values updated in '{fmap['column']}'")

    # STEP 7: Save output
    print_header("STEP 7: SAVE OUTPUT FILE")
    df.drop(columns=[age_cfg["age_variable"]], errors='ignore').to_excel(config["output_file"], index=False)
    print(f"✅ File saved: {config['output_file']}")

    # Done
    print_header("✅ PROCESS COMPLETE")
    print(f"⏱️ Time taken: {time.time() - start_time:.2f} seconds")

# === RUN ===
if __name__ == "__main__":
    process_excel(CONFIG)
