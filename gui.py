import tkinter as tk
from tkinter import ttk, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import pandas as pd
from lifelines import KaplanMeierFitter

from cli import analyze, EXCEL_FILE


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
    tree["columns"] = list(df.columns)
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, anchor=tk.CENTER, width=120)
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))


def main() -> None:
    pred_df = analyze() or pd.DataFrame()
    events_df = load_events()

    root = tk.Tk()
    root.title("Cable Predictor")

    tree = ttk.Treeview(root, show="headings")
    tree.pack(fill=tk.BOTH, expand=True)
    if not pred_df.empty:
        populate_tree(tree, pred_df)

    control_frame = ttk.Frame(root)
    control_frame.pack(fill=tk.X)

    connector_var = tk.StringVar()
    connector_box = ttk.Combobox(
        control_frame,
        textvariable=connector_var,
        values=sorted(events_df["connector_type"].unique().tolist()),
    )
    connector_box.pack(side=tk.LEFT, padx=5, pady=5)

    fig = Figure(figsize=(5, 4))
    ax = fig.add_subplot(111)
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def plot_curve() -> None:
        connector = connector_var.get()
        ax.clear()
        if not connector:
            canvas.draw()
            return
        sub = events_df[events_df["connector_type"] == connector]
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

    ttk.Button(control_frame, text="Plot", command=plot_curve).pack(
        side=tk.LEFT, padx=5, pady=5
    )

    root.mainloop()


if __name__ == "__main__":
    main()
