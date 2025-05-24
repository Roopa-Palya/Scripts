import pandas as pd
import os
from datetime import datetime

# === CONFIGURATION ===
CONFIG = {
    "input_excel": "input_data.xlsx",                    # Input Excel file name
    "output_excel": "output_with_new_columns.xlsx",      # Output Excel file name
    "new_columns": ["Scan Date", "Reviewer", "Remarks"]  # New columns to add at the beginning
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

    # Step 1: Check if input file exists
    print(f"🔍 Checking for file: '{CONFIG['input_excel']}' in current directory...")
    if not os.path.exists(CONFIG["input_excel"]):
        print(f"❌ File '{CONFIG['input_excel']}' not found. Please check the file name.")
        return
    print("✅ File found!")

    # Step 2: Read the input Excel file
    print("📖 Reading the Excel file...")
    df_original = pd.read_excel(CONFIG["input_excel"], engine='openpyxl')
    print(f"✅ Successfully read the file. Original columns are: {list(df_original.columns)}")

    # Step 3: Prepare new empty columns DataFrame
    print(f"➕ Adding new columns at the beginning: {CONFIG['new_columns']}")
    df_new = pd.DataFrame(columns=CONFIG["new_columns"])

    # Step 4: Combine new empty columns with original data
    df_combined = pd.concat([df_new, df_original], axis=1)
    print("🧩 Combined new columns with original data.")

    # Step 5: Write to output file
    print(f"💾 Writing to output file: '{CONFIG['output_excel']}'...")
    df_combined.to_excel(CONFIG["output_excel"], index=False)
    print("✅ File saved successfully!")

    # Step 6: Timer Summary
    end_time = datetime.now()
    duration = end_time - start_time
    print(f"🕒 Execution Time: {format_duration(duration)}\n")
    print("🎉 Script execution completed.\n")

if __name__ == "__main__":
    main()
