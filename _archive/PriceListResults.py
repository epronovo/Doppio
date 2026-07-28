import os
import csv
from collections import defaultdict

# Set the folder path containing .csv files
FOLDER_PATH = "./PriceLists"  # change this to your actual folder path
OUTPUT_FILE = r"C:/ASRaymond/PriceListResults.csv"

summary_data = []

# Iterate through all .csv files in the folder
for filename in os.listdir(FOLDER_PATH):
    if filename.endswith(".csv"):
        error_counter = defaultdict(int)
        file_path = os.path.join(FOLDER_PATH, filename)

        with open(file_path, "r", encoding="utf-8") as file:
            for line in file:
                if "already exists" in line:
                    continue
                if line.startswith("##"):
                    error_key = line[:14]  # First 14 characters
                    error_counter[error_key] += 1

        # Append results to summary_data
        for error, count in error_counter.items():
            summary_data.append([filename, error, count])

# Write OCOU_Summary to a new CSV file
with open(OUTPUT_FILE, "w", newline="", encoding="utf-8") as out_csv:
    writer = csv.writer(out_csv)
    writer.writerow(["file name", "error", "error count"])
    writer.writerows(summary_data)

print(f"OCOU_Summary written to {OUTPUT_FILE}")