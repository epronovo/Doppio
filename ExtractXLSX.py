import time
import pandas as pd
import sqlite3
import os
from InforMI import prompt_for_company_division
from config import BASE_DIR, get_sqlite_db_path
import re

# initialize field
print("")
QUERIES_DIR = BASE_DIR / "queries/xlsx"
prompt_for_company_division()
SQLITE_DB_PATH = get_sqlite_db_path()

def load_excel_files():
    excel_files = {}
    for filename in os.listdir(QUERIES_DIR):
        if filename.endswith(".xlsx"):
            filepath = os.path.join(QUERIES_DIR, filename)
            excel_files[filename] = filepath
    return excel_files

def recreate_sqlite_table(table_name, df):
    conn = sqlite3.connect(SQLITE_DB_PATH)
    df.columns = [col.strip().replace(" ", "_") for col in df.columns]
    # Convert float columns to string to avoid precision issues
    for col in df.select_dtypes(include=['float']).columns:
        def format_float(x):
            if pd.isnull(x):
                return ""
            elif x == int(x):
                return str(int(x))  # whole number as text
            else:
                return f"{x:.4f}"   # float as text
        df[col] = df[col].apply(format_float)
    df.astype(str).to_sql(table_name, conn, if_exists='replace', index=False)
    conn.close()
def get_workbook_prefix(filename):
    # Remove extension
    name_no_ext = os.path.splitext(filename)[0]
    # Extract capital letters (CamelCase or PascalCase words)
    prefix = "".join(re.findall(r'[A-Z]', name_no_ext))
    return prefix or name_no_ext[:3].upper()  # fallback if no capitals found

if __name__ == "__main__":
    start_time = time.time()  # Start timer
    files = load_excel_files()

    for filename, filepath in files.items():
        print(f"Processing: {filename}")
        prefix = get_workbook_prefix(filename)

        try:
            excel_data = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")
            for sheet_name, df in excel_data.items():
                table_name = f"{prefix}_{sheet_name.strip().replace(' ', '_')}"
                recreate_sqlite_table(table_name, df)
                print(f"Loaded sheet '{sheet_name}' into table '{table_name}'")
        except Exception as e:
            print(f"Failed to process {filename}: {e}")

    elapsed = time.time() - start_time  # End timer
    hours, rem = divmod(elapsed, 3600)
    minutes, seconds = divmod(rem, 60)
    print(f"Total run time: {int(hours):02d}:{int(minutes):02d}:{int(seconds):02d} (hh:mm:ss)")