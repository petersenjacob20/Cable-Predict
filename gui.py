import io
import tkinter as tk
from tkinter import ttk, messagebox
from contextlib import redirect_stdout

import pandas as pd
from lifelines import KaplanMeierFitter
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from cli import add_event, analyze, EXCEL_FILE


def load_events() -> pd.DataFrame:
    """Load the Events sheet from the Excel file."""
    if EXCEL_FILE.exists():
        try:
            return pd.read_excel(EXCEL_FILE, sheet_name="Events")
        except ValueError:
            pass
    return pd.DataFrame(columns=["connector_type", "serial_number", "cycles", "event"])


def populate_tree(tree: ttk.Treeview, df: pd.DataFrame) -> None:
    """Populate a Treeview with a DataFrame."""
    tree.delete(*tree.get_children())
    if df.empty:
        tree["columns"] = []
        return
    cols = list(df.columns)
    tree["columns"] = cols
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, anchor=tk.CENTER, width=120)
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))


def make_gui() -> None:
    root = tk.Tk()
    root.title("Cable Predictor")

    # ---------- Top: Add Event ----------
    add_frame = ttk.LabelFrame(root, text="Add Event")
    add_frame.pack(fill="x", padx=10, pady=6)

    ttk.Label(add_frame, text="Connector Type:").grid(column=0, row=0, sticky="w", padx=5, pady=2)
    connector_entry = ttk.Entry(add_frame, width=24)
    connector_entry.grid(column=1, row=0, padx=5, pady=2)

    ttk.Label(add_frame, text="Serial Number:").grid(column=0, row=1, sticky="w", padx=5, pady=2)
    serial_entry = ttk.Entry(add_frame, width=24)
    serial_entry.grid(column=1, row=1, padx=5, pady=2)

    ttk.Label(add_frame, text="Cycles:").grid(column=0, row=2, sticky="w", padx=5, pady=2)
    cycles_entry = ttk.Entry(add_frame, width=24)
    cycles_entry.grid(column=1, row=2, padx=5, pady=2)

    add_msg = tk.StringVar()

    def submit_event() -> None:
        connector = connector_entry.get().strip()
        serial = serial_entry.get().strip()
        cycles_text = cycles_entry.get().strip()
        if not connector or not serial or not cycles_text:
            add_msg.set("All fields are required.")
            return
        try:
            cycles = int(cycles_text)
        except ValueError:
            add_msg.set("Cycles must be an integer.")
            return
        try:
            add_event(connector, serial, cycles)
            add_msg.set("Event added successfully.")
            refresh_data()
        except Exception as exc:
            add_msg.set(f"Failed to add event: {exc}")

    ttk.Button(add_frame, text="Submit Event", command=submit_event).grid(column=0, row=3, columnspan=2, pady=6)
    ttk.Label(add_frame, textvariable=add_msg).grid(column=0, row=4, columnspan=2, padx=5, pady=2)

    # ---------- Middle: Predictions table + controls ----------
    mid_frame = ttk.Frame(root)
    mid_frame.pack(fill="both", expand=True, padx=10, pady=6)

    tree = ttk.Treeview(mid_frame, show="headings", height=8)
    tree.pack(fill="both", expand=True)

    ctrl = ttk.Frame(root)
    ctrl.pack(fill="x", padx=10, pady=6)

    ttk.Label(ctrl, text="Connector:").pack(side=tk.LEFT, padx=5)
    connector_var = tk.StringVar()
    connector_box = ttk.Combobox(ctrl, textvariable=connector_var, width=24)
    connector_box.pack(side=tk.LEFT, padx=5)

    # ---------- Bottom: Plot ----------
    fig = Figure(figsize=(6, 3.5))
    ax = fig.add_subplot(111)
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.get_tk_widget().pack(fill="both", expand=True, padx=10, pady=6)

    # ---------- Actions ----------
    analysis_msg = tk.StringVar()

    def run_analysis() -> None:
        buf = io.StringIO()
        try:
            with redirect_stdout(buf):  # capture prints from analyze()
                df = analyze()  # returns DataFrame or None
            output = buf.getvalue().strip()
            analysis_msg.set(output or "Analysis completed.")
            refresh_table(df)
            refresh_connector_list()
        except Exception as exc:
            analysis_msg.set(f"Analysis failed: {exc}")

    def plot_curve() -> None:
        connector = connector_var.get().strip()
        ax.clear()
        ev = load_events()
        if not connector:
            canvas.draw()
            return
        sub = ev[ev["connector_type"] == connector]
        if sub.empty:
            messagebox.showinfo("No data", f"No events for {connector}")
            canvas.draw()
            return
        kmf = KaplanMeierFitter()
        kmf.fit(sub["cycles"], event_observed=sub["event"], label=connector)
        kmf.plot_survival_function(ax=ax)
        ax.set_xlabel("Cycles")
        ax.set_ylabel("Survival probability")
        ax.set_title(f"Survival Curve - {connector}")
        canvas.draw()

    ttk.Button(ctrl, text="Run Analysis", command=run_analysis).pack(side=tk.LEFT, padx=5)
    ttk.Button(ctrl, text="Plot", command=plot_curve).pack(side=tk.LEFT, padx=5)
    ttk.Label(ctrl, textvariable=analysis_msg).pack(side=tk.LEFT, padx=10)

    # ---------- Helpers to refresh UI ----------
    def refresh_table(df: pd.DataFrame | None) -> None:
        if df is None or df.empty:
            # try loading from Excel if analyze returned None
            try:
                df = pd.read_excel(EXCEL_FILE, sheet_name="Predictions")
            except Exception:
                df = pd.DataFrame()
        populate_tree(tree, df)

    def refresh_connector_list() -> None:
        ev = load_events()
        connector_box["values"] = sorted(ev["connector_type"].dropna().unique().tolist())

    def refresh_data() -> None:
        refresh_table(None)
        refresh_connector_list()

    # initial load
    refresh_data()

    root.mainloop()


if __name__ == "__main__":
    make_gui()
