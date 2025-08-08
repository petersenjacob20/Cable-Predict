# main.py
from uploadTxtFiles import parse_log_file, is_already_logged
from updateSpreadsheet import update_cable_counts
from helper import ensure_excel_with_sheets

import os

LOG_FOLDER = "./logs"
EXCEL_FILE = "Cable Tracker.xlsx"

def main():
    ensure_excel_with_sheets(EXCEL_FILE)
    for filename in os.listdir(LOG_FOLDER):
        if filename.endswith(".log") or filename.endswith(".txt"):
            filepath = os.path.join(LOG_FOLDER, filename)
            data = parse_log_file(filepath)
            if data and not is_already_logged(EXCEL_FILE, data["timestamp"]):
                update_cable_counts(EXCEL_FILE, data)
            else:
                print(f"Skipping {filename}: Already logged or not ATP test.")

if __name__ == "__main__":
    main()
