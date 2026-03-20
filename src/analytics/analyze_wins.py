import pandas as pd
import tkinter as tk
from tkinter import filedialog


EXCLUDED_STATS = {"P", "PTS", "PPP", "PKP"}
STAT_COEFFICIENT_ALIASES = {
    "G": "A",
    "PPG": "PPA",
    "PKG": "PKA",
}


def calculate_coefficients(data, positive_stats, negative_stats, win_columns):
    coefficients = {}
    for stat in positive_stats + negative_stats:
        if stat.upper() in EXCLUDED_STATS:
            continue
        if stat not in data.columns:
            print(f"Warning: Stat column '{stat}' not found in data.")
            continue
        if data[stat].sum() == 0:
            print(f"Skipping '{stat}' as all values are 0.")
            continue
        for win_col in win_columns:
            if win_col not in data.columns:
                print(f"Warning: Win column '{win_col}' not found in data.")
                continue
            try:
                correlation = data[stat].corr(data[win_col])
                if pd.isna(correlation):
                    print(f"Correlation between '{stat}' and '{win_col}' is NaN.")
                    continue
                if stat in positive_stats:
                    coefficients[f"{stat} vs {win_col}"] = max(0, correlation)
                elif stat in negative_stats:
                    coefficients[f"{stat} vs {win_col}"] = -abs(correlation)
            except Exception as e:
                print(
                    "Error calculating correlation between '"
                    f"{stat}' and '{win_col}': {e}"
                )
    if not coefficients:
        return None

    win_suffixes = {key.split(" vs ", 1)[1] for key in coefficients if " vs " in key}
    for target, source in STAT_COEFFICIENT_ALIASES.items():
        for win_suffix in win_suffixes:
            source_key = f"{source} vs {win_suffix}"
            if source_key in coefficients:
                coefficients[f"{target} vs {win_suffix}"] = coefficients[source_key]

    return coefficients if coefficients else None


def analyze_wins(vhl_regular, vhl_playoff, vhlm_regular, vhlm_playoff, output_file):
    # Reintroduce stat definitions
    positive_skater_stats = [
        "G",
        "A",
        "+/-",
        "HIT",
        "SHT",
        "SB",
        "PPG",
        "PPA",
        "PKG",
        "PKA",
        "SCHT",
        "TA",
        "PRET",
        "PI",
    ]
    negative_skater_stats = ["PIM", "HTT", "SCHTA", "GA", "PIA"]

    positive_goalie_stats = [
        "PCT",
        "GA",
        "SA",
        "SAR",
        "PS%",
        "PSA",
    ]
    negative_goalie_stats = ["PIM"]

    win_columns = ["W"]

    # Combine all datasets
    combined_skaters = pd.concat(
        [
            pd.read_excel(vhl_regular, sheet_name="Skaters - All Teams"),
            pd.read_excel(vhl_playoff, sheet_name="Skaters - All Teams"),
            pd.read_excel(vhlm_regular, sheet_name="Skaters - All Teams"),
            pd.read_excel(vhlm_playoff, sheet_name="Skaters - All Teams"),
        ],
        ignore_index=True,
    )

    combined_goalies = pd.concat(
        [
            pd.read_excel(vhl_regular, sheet_name="Goalies - All Teams"),
            pd.read_excel(vhl_playoff, sheet_name="Goalies - All Teams"),
            pd.read_excel(vhlm_regular, sheet_name="Goalies - All Teams"),
            pd.read_excel(vhlm_playoff, sheet_name="Goalies - All Teams"),
        ],
        ignore_index=True,
    )

    # Sort coefficients by value in descending order
    def sort_coefficients(coefficients):
        return {
            k: v
            for k, v in sorted(
                coefficients.items(), key=lambda item: item[1], reverse=True
            )
        }

    # Calculate coefficients for combined datasets
    results = {}
    print("Processing Combined Skaters...")
    skater_coefficients = calculate_coefficients(
        combined_skaters,
        positive_skater_stats,
        negative_skater_stats,
        win_columns,
    )
    if skater_coefficients:
        skater_coefficients = sort_coefficients(skater_coefficients)
        results["Combined Skaters"] = skater_coefficients
    else:
        print("No valid coefficients calculated for Combined Skaters.")

    print("Processing Combined Goalies...")
    goalie_coefficients = calculate_coefficients(
        combined_goalies,
        positive_goalie_stats,
        negative_goalie_stats,
        win_columns,
    )
    if goalie_coefficients:
        goalie_coefficients = sort_coefficients(goalie_coefficients)
        results["Combined Goalies"] = goalie_coefficients
    else:
        print("No valid coefficients calculated for Combined Goalies.")

    # Write results to an Excel file
    with pd.ExcelWriter(output_file) as writer:
        for key, value in results.items():
            if value:
                df = pd.DataFrame(
                    list(value.items()), columns=["Comparison", "Coefficient"]
                )
                df.to_excel(writer, sheet_name=key, index=False)
            else:
                print(f"Skipping {key} due to no valid data.")


if __name__ == "__main__":
    # Create a file dialog for input selection
    root = tk.Tk()
    root.withdraw()

    print("Select VHL Regular Season file")
    vhl_regular = filedialog.askopenfilename(title="Select VHL Regular Season file")

    print("Select VHL Playoff file")
    vhl_playoff = filedialog.askopenfilename(title="Select VHL Playoff file")

    print("Select VHLM Regular Season file")
    vhlm_regular = filedialog.askopenfilename(title="Select VHLM Regular Season file")

    print("Select VHLM Playoff file")
    vhlm_playoff = filedialog.askopenfilename(title="Select VHLM Playoff file")

    output_file = "Results/Wins_Coefficient.xlsx"

    analyze_wins(
        vhl_regular,
        vhl_playoff,
        vhlm_regular,
        vhlm_playoff,
        output_file,
    )

    print(f"Analysis complete. Results saved to {output_file}")
