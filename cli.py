import argparse
from pathlib import Path
from typing import Optional

import pandas as pd
from lifelines import KaplanMeierFitter
from openpyxl import Workbook, load_workbook

EXCEL_FILE = Path(__file__).resolve().parent / "Cable Tester Count" / "Cable Tracker.xlsx"


def add_event(connector_type: str, serial_number: str, cycles: int) -> None:
    """Append a failure event to the Events sheet."""
    EXCEL_FILE.parent.mkdir(parents=True, exist_ok=True)
    if EXCEL_FILE.exists():
        wb = load_workbook(EXCEL_FILE)
    else:
        wb = Workbook()
    if "Events" in wb.sheetnames:
        ws = wb["Events"]
    else:
        ws = wb.create_sheet("Events")
        ws.append(["connector_type", "serial_number", "cycles", "event"])
    ws.append([connector_type, serial_number, cycles, 1])
    wb.save(EXCEL_FILE)


def analyze() -> Optional[pd.DataFrame]:
    """Perform survival analysis and write predictions.

    Returns the predictions DataFrame or ``None`` if analysis could not be
    performed.
    """
    if not EXCEL_FILE.exists():
        print("Excel file not found.")
        return None
    try:
        events_df = pd.read_excel(EXCEL_FILE, sheet_name="Events")
    except ValueError:
        print("Events sheet not found.")
        return None
    results = []
    for connector, group in events_df.groupby("connector_type"):
        kmf = KaplanMeierFitter()
        kmf.fit(group["cycles"], event_observed=group["event"])
        sf = kmf.survival_function_
        cycles80 = sf[sf["KM_estimate"] <= 0.8].index.min()
        cycles90 = sf[sf["KM_estimate"] <= 0.9].index.min()
        results.append(
            {
                "connector_type": connector,
                "median_cycles": kmf.median_survival_time_,
                "cycles_80_survival": float(cycles80) if pd.notna(cycles80) else None,
                "cycles_90_survival": float(cycles90) if pd.notna(cycles90) else None,
            }
        )
    pred_df = pd.DataFrame(results)
    wb = load_workbook(EXCEL_FILE)
    if "Predictions" in wb.sheetnames:
        del wb["Predictions"]
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a") as writer:
        writer.book = wb
        writer.sheets = {ws.title: ws for ws in wb.worksheets}
        pred_df.to_excel(writer, sheet_name="Predictions", index=False)
    print("Analysis written to Predictions sheet.")
    return pred_df


def main() -> None:
    parser = argparse.ArgumentParser(description="Cable tracker CLI")
    subparsers = parser.add_subparsers(dest="command")

    add_parser = subparsers.add_parser("add-event", help="Add a cable event")
    add_parser.add_argument("--connector", required=True, help="Connector type")
    add_parser.add_argument("--serial", required=True, help="Serial number")
    add_parser.add_argument("--cycles", required=True, type=int, help="Cycle count")

    subparsers.add_parser("analyze", help="Run survival analysis")

    args = parser.parse_args()
    if args.command == "add-event":
        add_event(args.connector, args.serial, args.cycles)
    elif args.command == "analyze":
        analyze()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
