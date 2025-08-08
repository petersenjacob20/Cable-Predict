import io
import tkinter as tk
from tkinter import ttk
from contextlib import redirect_stdout

from cli import add_event, analyze


def submit_event() -> None:
    connector = connector_entry.get().strip()
    serial = serial_entry.get().strip()
    cycles_text = cycles_entry.get().strip()
    if not connector or not serial or not cycles_text:
        add_message.set("All fields are required.")
        return
    try:
        cycles = int(cycles_text)
        add_event(connector, serial, cycles)
        add_message.set("Event added successfully.")
    except Exception as exc:  # pragma: no cover - GUI feedback
        add_message.set(f"Failed to add event: {exc}")


def run_analysis() -> None:
    buf = io.StringIO()
    try:
        with redirect_stdout(buf):
            analyze()
        output = buf.getvalue().strip()
        analysis_message.set(output or "Analysis completed.")
    except Exception as exc:  # pragma: no cover - GUI feedback
        analysis_message.set(f"Analysis failed: {exc}")


root = tk.Tk()
root.title("Cable Predictor")

# Add Event Section
add_frame = ttk.LabelFrame(root, text="Add Event")
add_frame.pack(fill="x", padx=10, pady=5)

ttk.Label(add_frame, text="Connector Type:").grid(column=0, row=0, sticky="w", padx=5, pady=2)
connector_entry = ttk.Entry(add_frame)
connector_entry.grid(column=1, row=0, padx=5, pady=2)

ttk.Label(add_frame, text="Serial Number:").grid(column=0, row=1, sticky="w", padx=5, pady=2)
serial_entry = ttk.Entry(add_frame)
serial_entry.grid(column=1, row=1, padx=5, pady=2)

ttk.Label(add_frame, text="Cycles:").grid(column=0, row=2, sticky="w", padx=5, pady=2)
cycles_entry = ttk.Entry(add_frame)
cycles_entry.grid(column=1, row=2, padx=5, pady=2)

add_message = tk.StringVar()
ttk.Button(add_frame, text="Submit Event", command=submit_event).grid(column=0, row=3, columnspan=2, pady=5)
ttk.Label(add_frame, textvariable=add_message).grid(column=0, row=4, columnspan=2, padx=5, pady=2)

# Run Analysis Section
analysis_frame = ttk.LabelFrame(root, text="Run Analysis")
analysis_frame.pack(fill="x", padx=10, pady=5)

analysis_message = tk.StringVar()
ttk.Button(analysis_frame, text="Run Analysis", command=run_analysis).pack(padx=5, pady=5)
ttk.Label(analysis_frame, textvariable=analysis_message, wraplength=400).pack(padx=5, pady=2)

root.mainloop()
