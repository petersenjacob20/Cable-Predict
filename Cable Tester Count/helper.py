"""Utility helpers for the Cable Tester scripts."""

import os
from openpyxl import Workbook


def format_timestamp(ts_string):
    """Convert a timestamp string into a :class:`datetime` object."""
    from datetime import datetime

    return datetime.strptime(ts_string, "%Y-%m-%d %H:%M:%S")


def ensure_excel_with_sheets(excel_file: str) -> None:
    """Ensure an Excel workbook exists.

    Creates a blank workbook at ``excel_file`` if the path does not exist.
    The generated workbook contains the default sheet provided by
    :mod:`openpyxl`. Subsequent functions are responsible for creating any
    additional sheets as needed.
    """

    if not os.path.exists(excel_file):
        wb = Workbook()
        wb.save(excel_file)
        wb.close()
