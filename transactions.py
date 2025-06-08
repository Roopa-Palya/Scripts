import pandas as pd
import os
import time
from datetime import datetime

def print_header(title):
    print(f"\n{'=' * 60}\n🔷 {title}\n{'=' * 60}")

def process_excel(config):
    start_time = time.time()

    # STEP 1: Load Excel File
    print_header("STEP 1: LOAD EXCEL FILE")
    if not os.path.exists(config["input_file"]):
        print(f"❌ File not found: {config['input_file']}")
        return
    df = pd.read_excel(config["input_file"], sheet_name=config.get("input_sheet", 0))
    print(f"✅ Loaded {df.shape[0]} rows from {config['input_file']}")

    # STEP 2: Keep Only Specified Columns
    print_header("STEP 2: SELECT COLUMNS")
    df = df[[col for col in config["columns_to_extract"] if col in df.columns]]
    if config.get("drop_empty_rows"):
        before = df.shape[0]
        df = df.dropna(how="all")
        print(f"🧹 Dropped {before - df.shape[0]} empty rows.")

    # STEP 3: Add Static/Blank Columns
    print_header("STEP 3: ADD NEW COLUMNS")
    for col, val in config["columns_to_add"].items():
        df[col] = val
        print(f"➕ Added column: {col} = '{val}'")

    # STEP 4: Calculate Age (days)
    print_header("STEP 4: CALCULATE AGE")
    date_col = config["derived_column"]["based_on_column"]
    age_col = config["derived_column"]["column_name"]
    today = datetime.today().date()

    df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.date
    df[age_col] = df[date_col].apply(lambda d: (today - d).days if pd.notnull(d) else "")
    print(f"🧮 Age calculated in column '{age_col}' from '{date_col}'")

    # STEP 5: Apply EOL or SLO Rules into 'Lifecycle'
    print_header("STEP 5: APPLY LIFECYCLE + SLO RULES")
    lconf = config["lifecycle_rule"]
    slo_rules = config["slo_rules"]

    def evaluate_lifecycle(row):
        # EOL check
        if lconf["keyword"].lower() in str(row.get(lconf["trigger_column"], "")).lower():
            return lconf["value"]

        # SLO check
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
    print(f"✅ 'Lifecycle' column filled using both EOL and SLO rules.")

    # STEP 6: Save Output
    print_header("STEP 6: SAVE OUTPUT FILE")
    df.to_excel(config["output_file"], index=False)
    print(f"✅ Output written to '{config['output_file']}'")

    # Summary
    print_header("✅ PROCESS COMPLETE")
    print(f"⏱️ Duration: {time.time() - start_time:.2f} seconds")

# === CONFIG + RUN ===
if __name__ == "__main__":
    CONFIG = {  # full config from previous block pasted here }
    process_excel(CONFIG)
