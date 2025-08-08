from survivalAnalyzer import run_survival_analysis
import matplotlib.pyplot as plt

def main():
    excel_file = "../Cable Tester Count/Cable Tracker.xlsx"

    results, survival_curves = run_survival_analysis(excel_file)
    
    # Table
    print("\nRecommended Replacement Points:")
    print(f"{'Connector Type':<20} {'Replace After X Uses':<25}")
    print("-" * 45)
    for key, (x, y) in survival_curves.items():
        x80 = next((xi for xi, yi in zip(x, y) if yi <= 0.2), None)
        print(f"{key:<80} {x80 if x80 else 'N/A':<25}")

    # Plot
    plt.figure(figsize=(10, 6))
    for key, (x, y) in survival_curves.items():
        plt.plot(x, y, label=key)
    plt.axhline(y=0.2, color='gray', linestyle='--', label='20% Threshold')
    plt.title("Survival Curve per Connector Type")
    plt.xlabel("Number of Uses")
    plt.ylabel("Probability of Still Working")
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.savefig("survival_curve.png")
    plt.show()


if __name__ == "__main__":
    main()
