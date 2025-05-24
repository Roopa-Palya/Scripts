import pandas as pd
import os
from datetime import datetime

# === CONFIGURATION ===
CONFIG = {
    "input_excel": "input_data.xlsx",                    # Input Excel file name
    "output_excel": "output_with_new_columns.xlsx",      # Output Excel file name
    "new_columns": ["Scan Date", "Reviewer", "Remarks"], # New columns to add
    "column_static_values": {                            # Static values for new columns
        "Scan Date": "2025-05-24",
        "Reviewer": "Security Team"
        # "Remarks" will be left blank
    }
}

def format_duration(duration):
    seconds = duration.total_seconds()
    if seconds < 60:
        return f"{seconds:.2f} seconds"
    elif seconds < 3600:
        return f"{seconds // 60:.0f} minutes {seconds % 60:.0f} seconds"
    else:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        seconds = int(seconds % 60)
        return f"{hours} hours {minutes} minutes {seconds} seconds"

def main():
    print("\n🚀 Starting the Excel column processing script...")
    start_time = datetime.now()

    print(f"🔍 Checking for file: '{CONFIG['input_excel']}' in current directory...")
    if not os.path.exists(CONFIG["input_excel"]):
        print(f"❌ File '{CONFIG['input_excel']}' not found. Please check the file name.")
        return
    print("✅ File found!")

    print("📖 Reading the Excel file...")
    df_original = pd.read_excel(CONFIG["input_excel"], engine='openpyxl')
    print(f"✅ Successfully read the file. Original columns: {list(df_original.columns)}")

    print(f"➕ Adding new columns: {CONFIG['new_columns']}")
    df_new = pd.DataFrame()

    for col in CONFIG["new_columns"]:
        if col in CONFIG.get("column_static_values", {}):
            static_value = CONFIG["column_static_values"][col]
            df_new[col] = [static_value] * len(df_original)
            print(f"🧷 Assigned static value '{static_value}' to column '{col}'")
        else:
            df_new[col] = ["" for _ in range(len(df_original))]
            print(f"⬜ Column '{col}' will be left empty")

    df_combined = pd.concat([df_new, df_original], axis=1)
    print("🧩 Combined new columns with original data.")

    print(f"💾 Writing to output file: '{CONFIG['output_excel']}'...")
    df_combined.to_excel(CONFIG["output_excel"], index=False)
    print("✅ File saved successfully!")

    end_time = datetime.now()
    print(f"🕒 Execution Time: {format_duration(end_time - start_time)}")
    print("🎉 Script execution completed.\n")

if __name__ == "__main__":
    main()
