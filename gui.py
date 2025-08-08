import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from typing import Optional

from cli import add_event, analyze


class CableTrackerGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Cable Tracker")
        self.excel_file: Optional[Path] = None

        tk.Label(root, text="Connector Type").grid(row=0, column=0, sticky="e")
        self.connector_entry = tk.Entry(root)
        self.connector_entry.grid(row=0, column=1)

        tk.Label(root, text="Serial Number").grid(row=1, column=0, sticky="e")
        self.serial_entry = tk.Entry(root)
        self.serial_entry.grid(row=1, column=1)

        tk.Label(root, text="Cycles").grid(row=2, column=0, sticky="e")
        self.cycles_entry = tk.Entry(root)
        self.cycles_entry.grid(row=2, column=1)

        tk.Button(root, text="Select Workbook", command=self.select_workbook).grid(
            row=3, column=0, columnspan=2, pady=(5, 0)
        )
        tk.Button(root, text="Add Event", command=self.add_event).grid(
            row=4, column=0, columnspan=2, pady=(5, 0)
        )
        tk.Button(root, text="Analyze", command=self.run_analyze).grid(
            row=5, column=0, columnspan=2, pady=(5, 0)
        )

    def select_workbook(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select workbook", filetypes=[("Excel Files", "*.xlsx")]
        )
        if file_path:
            self.excel_file = Path(file_path)
        else:
            self.excel_file = None

    def add_event(self) -> None:
        if not self.excel_file:
            messagebox.showerror("Error", "Please select a workbook first.")
            return
        connector = self.connector_entry.get().strip()
        serial = self.serial_entry.get().strip()
        try:
            cycles = int(self.cycles_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Cycles must be an integer.")
            return
        add_event(self.excel_file, connector, serial, cycles)
        messagebox.showinfo("Success", "Event added.")

    def run_analyze(self) -> None:
        if not self.excel_file:
            messagebox.showerror("Error", "Please select a workbook first.")
            return
        analyze(self.excel_file)
        messagebox.showinfo("Success", "Analysis complete.")


if __name__ == "__main__":
    root = tk.Tk()
    gui = CableTrackerGUI(root)
    root.mainloop()
