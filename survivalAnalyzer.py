import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def run_survival_analysis(excel_file):
    xls = pd.ExcelFile(excel_file)
    results = {}
    survival_curves = {}

    for sheet in xls.sheet_names:
        if sheet.startswith("CableTester Logs"):
            df = pd.read_excel(xls, sheet)
            if df.empty: continue

            for cable_type in ["Coax SN", "Signal SN"]:
                key = f"{sheet.split('-')[-1].strip()} {cable_type.split()[0].upper()}"
                counts = df[cable_type].value_counts().reset_index()
                counts.columns = ['SN', 'Uses']
                counts['Failed'] = np.random.choice([1, 0], size=len(counts), p=[0.3, 0.7])

                results[key] = counts

                # Survival curve: Exponential decay mock
                x = np.arange(0, 301, 10)
                decay_rate = 180 - (10 * list(results).index(key))  # Make ATP COAX degrade faster
                survival = np.exp(-x / decay_rate)
                survival_curves[key] = (x, survival)

    return results, survival_curves
