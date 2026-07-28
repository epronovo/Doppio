import os
import pandas as pd

# Folder containing your Excel files
FOLDER_PATH = "./evs100/Complete"  # change this to your folder path
OUTPUT_FILE = "PriceListSummary.report.txt"

def collect_error_summary(folder_path):
    summary_set = set()

    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx") and "_e" in filename:
            filepath = os.path.join(folder_path, filename)
            print(f"Processing: {filename}")
            
            # Read all sheets
            xls = pd.ExcelFile(filepath)
            for sheet_name in xls.sheet_names:
                try:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    if {'MESSAGE', 'ITNO'}.issubset(df.columns):
                        for _, row in df.iterrows():
                            summary_set.add((str(row['MESSAGE']), str(row['ITNO'])))
                except Exception as e:
                    print(f"  Skipped sheet {sheet_name} due to error: {e}")

    return summary_set

if __name__ == "__main__":
    summary = collect_error_summary(FOLDER_PATH)
    
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        for message, itno in sorted(summary):
            f.write(f"MESSAGE: {message} | ITNO: {itno}\n")

    print(f"Summary written to {OUTPUT_FILE}")