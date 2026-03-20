import math
import os
import tkinter as tk
from tkinter import filedialog

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


SKATER_SHEET_KEYWORDS = ("skater",)
GOALIE_SHEET_KEYWORDS = ("goalie",)
SKATER_COEFF_SHEET_KEYWORDS = ("skater",)
GOALIE_COEFF_SHEET_KEYWORDS = ("goalie",)
BASELINE_RATIO = 1e-6
CATEGORY_WEIGHT_BLEND = 0.35
CATEGORY_WEIGHT_EXP = 0.75
PLAYER_RATIO_EXP = 0.75
TEAM_UNIFORM_BLEND = 0.4
CATEGORY_PLAYER_POWER = 0.5
EXCLUDED_STATS = {"P", "POINTS", "PTS", "PPP", "PKP"}
STAT_COEFFICIENT_ALIASES = {
    "G": "A",
    "PPG": "PPA",
    "PKG": "PKA",
}

POINT_TOTAL_WEIGHT = 4.0
POINT_GOAL_WEIGHT = 2.0
POINT_ASSIST_WEIGHT = 1.5
POINT_SHARE_EXP = 1.05
POINT_TOTAL_EXP = 1.1
POINT_PRIORITY = 0.7
POINT_SCALE_FLOOR = 1e-6

COLUMN_ALIASES = {
    "Team Name": "Team",
    "Player Name": "Player",
    "Goalie Name": "Player",
}


def _find_sheet(sheet_names, keywords):
    for sheet in sheet_names:
        lowered = sheet.lower()
        if any(keyword in lowered for keyword in keywords):
            return sheet
    return None


def _load_category_tables(excel_path):
    workbook = pd.ExcelFile(excel_path)
    sheets = workbook.sheet_names

    skater_sheet = _find_sheet(sheets, SKATER_SHEET_KEYWORDS)
    goalie_sheet = _find_sheet(sheets, GOALIE_SHEET_KEYWORDS)

    tables = {}
    if skater_sheet:
        tables["Skater"] = workbook.parse(skater_sheet)
    if goalie_sheet:
        tables["Goalie"] = workbook.parse(goalie_sheet)

    if not tables:
        # Fallback to the first sheet if no recognised names exist
        tables["Skater"] = workbook.parse(sheets[0])

    return tables


def _load_coefficients(excel_path):
    workbook = pd.ExcelFile(excel_path)
    sheets = workbook.sheet_names

    coeff_tables = {}
    skater_sheet = _find_sheet(sheets, SKATER_COEFF_SHEET_KEYWORDS)
    goalie_sheet = _find_sheet(sheets, GOALIE_COEFF_SHEET_KEYWORDS)

    if skater_sheet:
        coeff_tables["Skater"] = workbook.parse(skater_sheet)
    if goalie_sheet:
        coeff_tables["Goalie"] = workbook.parse(goalie_sheet)

    if not coeff_tables:
        coeff_tables["Skater"] = workbook.parse(sheets[0])

    return coeff_tables


def _build_coefficient_dict(coeff_df):
    mapping = {}
    for comparison, value in zip(coeff_df["Comparison"], coeff_df["Coefficient"]):
        stat_name = str(comparison).split(" vs ")[0].strip()
        if not stat_name:
            continue
        mapping[stat_name] = value

    filtered = {}
    for stat_name, value in mapping.items():
        if stat_name.upper() in EXCLUDED_STATS:
            continue
        filtered[stat_name] = value

    upper_lookup = {name.upper(): name for name in filtered}
    for target, source in STAT_COEFFICIENT_ALIASES.items():
        source_upper = source.upper()
        if source_upper in upper_lookup:
            filtered[target] = filtered[upper_lookup[source_upper]]
        elif source in filtered:
            filtered[target] = filtered[source]

    return filtered


def _harmonize_columns(df):
    for original, alias in COLUMN_ALIASES.items():
        if original in df.columns and alias not in df.columns:
            df[alias] = df[original]
    return df


def _adjust_scores(raw_scores):
    if not raw_scores:
        return {}, 0.0

    min_score = min(raw_scores.values())
    if min_score < 0:
        adjusted = {idx: score - min_score for idx, score in raw_scores.items()}
    else:
        adjusted = dict(raw_scores)

    adjusted = {idx: max(score, 0.0) for idx, score in adjusted.items()}
    total = sum(adjusted.values())

    if total <= 0:
        adjusted = {idx: 1.0 for idx in raw_scores}
        total = float(len(adjusted))
    else:
        baseline = total * BASELINE_RATIO if total else BASELINE_RATIO
        adjusted = {idx: score + baseline for idx, score in adjusted.items()}
        total = sum(adjusted.values())

    return adjusted, total


def _boost_point_contributions(raw_scores, team_df):
    if not raw_scores:
        return raw_scores

    has_goal_data = "G" in team_df.columns
    has_assist_data = "A" in team_df.columns
    has_point_data = "P" in team_df.columns

    if not (has_goal_data or has_assist_data or has_point_data):
        return raw_scores

    goal_series = team_df["G"].fillna(0.0) if has_goal_data else None
    assist_series = team_df["A"].fillna(0.0) if has_assist_data else None
    point_series = team_df["P"].fillna(0.0) if has_point_data else None

    goal_total = float(goal_series.sum()) if goal_series is not None else 0.0
    assist_total = float(assist_series.sum()) if assist_series is not None else 0.0
    if point_series is not None:
        point_total = float(point_series.sum())
    else:
        point_total = goal_total + assist_total

    if goal_total <= 0 and assist_total <= 0 and point_total <= 0:
        return raw_scores

    magnitude = max(abs(value) for value in raw_scores.values())
    base_scale = max(magnitude, POINT_SCALE_FLOOR)

    for idx, current_value in list(raw_scores.items()):
        goal_value = goal_series.at[idx] if goal_series is not None else 0.0
        assist_value = assist_series.at[idx] if assist_series is not None else 0.0
        if point_series is not None:
            point_value = point_series.at[idx]
        else:
            point_value = goal_value + assist_value

        goal_share = goal_value / goal_total if goal_total > 0 else 0.0
        assist_share = assist_value / assist_total if assist_total > 0 else 0.0
        point_share = point_value / point_total if point_total > 0 else 0.0

        goal_bonus = 0.0
        if goal_share > 0:
            goal_bonus = POINT_GOAL_WEIGHT * (goal_share**POINT_SHARE_EXP)

        assist_bonus = 0.0
        if assist_share > 0:
            assist_bonus = POINT_ASSIST_WEIGHT * (assist_share**POINT_SHARE_EXP)

        point_bonus = 0.0
        if point_share > 0:
            point_bonus = POINT_TOTAL_WEIGHT * (point_share**POINT_TOTAL_EXP)

        point_strength = base_scale * (goal_bonus + assist_bonus + point_bonus)
        blended_score = (
            1.0 - POINT_PRIORITY
        ) * current_value + POINT_PRIORITY * point_strength
        raw_scores[idx] = blended_score

    return raw_scores


def _move_column_after(df, column, after):
    if column in df.columns and after in df.columns:
        columns = df.columns.tolist()
        columns.remove(column)
        after_index = columns.index(after)
        columns.insert(after_index + 1, column)
        return df[columns]
    return df


def _build_combined_subset(df, category):
    columns = [
        "Player Name",
        "User",
        "Team Name",
        "POS",
        "GP",
        "WPA",
        "Win % Value",
        "Win %",
    ]
    if df.empty:
        return pd.DataFrame(columns=columns)

    if "Player Name" in df.columns:
        player_name = df["Player Name"].copy()
    elif "Player" in df.columns:
        player_name = df["Player"].copy()
    elif "Goalie Name" in df.columns:
        player_name = df["Goalie Name"].copy()
    else:
        player_name = pd.Series([None] * len(df), index=df.index)

    if "Team Name" in df.columns:
        team_name = df["Team Name"].copy()
    elif "Team" in df.columns:
        team_name = df["Team"].copy()
    else:
        team_name = pd.Series([None] * len(df), index=df.index)

    subset = pd.DataFrame(
        {
            "Player Name": player_name,
            "User": df.get("User"),
            "Team Name": team_name,
            "POS": df.get("POS"),
            "GP": df.get("GP"),
            "WPA": df.get("WPA"),
            "Win % Value": df.get("Win % Value"),
            "Win %": df.get("Win %"),
        }
    )

    return subset


def _normalize_percentage_shares(shares):
    if not shares:
        return []

    scaled_exact = [value * 100 for value in shares]
    floors = [int(math.floor(value + 1e-9)) for value in scaled_exact]
    total_basis = sum(floors)
    diff = int(round(10000 - total_basis))

    remainders = [value - floor for value, floor in zip(scaled_exact, floors)]
    if diff > 0:
        indices = sorted(
            range(len(shares)),
            key=lambda i: remainders[i],
            reverse=True,
        )
        for i in range(diff):
            floors[indices[i % len(indices)]] += 1
    elif diff < 0:
        indices = sorted(range(len(shares)), key=lambda i: remainders[i])
        for i in range(-diff):
            floors[indices[i % len(indices)]] -= 1

    return [value / 100 for value in floors]


def _attach_win_percentages(df, team_column, wpa_column, override_percentages=None):
    df = df.copy()

    percent_values = pd.Series(0.0, index=df.index)
    percent_strings = pd.Series("0.00%", index=df.index)

    if override_percentages is not None:
        if not isinstance(override_percentages, pd.Series):
            override_percentages = pd.Series(override_percentages)

        aligned = override_percentages.reindex(df.index).fillna(0.0)
        percent_values = aligned
        percent_strings = aligned.apply(lambda value: f"{value:.2f}%")
        df["Win % Value"] = percent_values
        df["Win %"] = percent_strings
        return df

    grouped = df.groupby(team_column, sort=False)
    for team, indices in grouped.groups.items():
        if pd.isna(team):
            continue

        team_indices = list(indices)
        team_total = df.loc[team_indices, wpa_column].sum()
        if team_total <= 0:
            continue

        raw_percentages = [
            (df.at[idx, wpa_column] / team_total) * 100 for idx in team_indices
        ]
        normalized = _normalize_percentage_shares(raw_percentages)
        for idx, value in zip(team_indices, normalized):
            percent_values.at[idx] = value
            percent_strings.at[idx] = f"{value:.2f}%"

    df["Win % Value"] = percent_values
    df["Win %"] = percent_strings
    return df


def _compute_team_win_percentages(categories, team_column, wpa_column):
    percent_lookup = {
        category: pd.Series(0.0, index=df.index) for category, df in categories.items()
    }

    team_entries = {}
    for category, df in categories.items():
        if team_column not in df.columns or wpa_column not in df.columns:
            continue

        for idx, row in df.iterrows():
            team = row.get(team_column)
            if pd.isna(team):
                continue

            wpa_value = row.get(wpa_column, 0.0)
            if pd.isna(wpa_value):
                continue

            contribution = max(float(wpa_value), 0.0)
            team_entries.setdefault(team, []).append((category, idx, contribution))

    for team, entries in team_entries.items():
        total = sum(value for _, _, value in entries)
        if total <= 0:
            continue

        raw_percentages = [
            (value / total) * 100 for _, _, value in entries if total > 0
        ]
        normalized = _normalize_percentage_shares(raw_percentages)

        for (category, idx, _), percent in zip(entries, normalized):
            percent_lookup[category].at[idx] = percent

    return percent_lookup


def _add_takeaway_difference(df):
    if "TA" not in df.columns or "GA" not in df.columns:
        return df

    ta_series = df["TA"].fillna(0.0)
    ga_series = df["GA"].fillna(0.0)
    tod_series = ta_series - ga_series

    if "TO" in df.columns:
        columns = list(df.columns)
        insert_index = columns.index("TO")
        df = df.drop(columns=["TO"])
        df.insert(insert_index, "TOD", tod_series)
    else:
        df["TOD"] = tod_series

    return df


def _auto_fit_columns(workbook_path):
    try:
        workbook = load_workbook(workbook_path)
    except Exception as exc:
        print(f"Unable to auto-fit columns for {workbook_path}: {exc}")
        return

    for sheet in workbook.worksheets:
        max_widths = {}
        for row in sheet.iter_rows(values_only=True):
            if row is None:
                continue
            for col_index, cell_value in enumerate(row, start=1):
                text = "" if cell_value is None else str(cell_value)
                existing = max_widths.get(col_index, 0)
                max_widths[col_index] = max(existing, len(text))

        for col_index in range(1, sheet.max_column + 1):
            header_value = sheet.cell(row=1, column=col_index).value
            header_text = "" if header_value is None else str(header_value)
            existing = max_widths.get(col_index, 0)
            max_widths[col_index] = max(existing, len(header_text))

        for col_index, width in max_widths.items():
            column_letter = get_column_letter(col_index)
            adjusted_width = min(max(width + 2, 8), 80)
            sheet.column_dimensions[column_letter].width = adjusted_width

    try:
        workbook.save(workbook_path)
    except Exception as exc:
        print(f"Unable to save auto-fitted workbook {workbook_path}: {exc}")


def _build_position_ranking(df, positions, avg_roster_size):
    if df.empty:
        return df

    pos_series = df["POS"].fillna("")
    mask = pos_series.str.upper().isin({pos.upper() for pos in positions})
    ranking_df = df.loc[mask].copy()

    if ranking_df.empty:
        return ranking_df

    roster_series = ranking_df.get("Roster Size")
    if roster_series is None:
        roster_series = pd.Series(avg_roster_size, index=ranking_df.index)

    roster_filled = roster_series.fillna(avg_roster_size).replace(0, avg_roster_size)
    ranking_df["Roster Size"] = roster_filled.apply(lambda value: int(round(value)))

    win_rate_series = ranking_df.get("Team Win % Value")
    if win_rate_series is None:
        win_rate_series = pd.Series(1.0, index=ranking_df.index)
    win_rate_series = win_rate_series.fillna(1.0)

    adjusted_values = ranking_df["Win % Value"] * win_rate_series
    ranking_df["Adjusted Win % Value"] = adjusted_values
    ranking_df["Adjusted Win %"] = ranking_df["Adjusted Win % Value"].apply(
        lambda value: f"{value:.2f}%"
    )

    ranking_df = ranking_df.sort_values(
        by=["Adjusted Win % Value", "Win % Value", "Player Name"],
        ascending=[False, False, True],
    )
    return ranking_df


def _determine_team_wins(team, categories, team_column, wins_column):
    skater_df = categories.get("Skater")
    if skater_df is not None:
        team_skater_df = skater_df[skater_df[team_column] == team]
        unique_wins = team_skater_df[wins_column].dropna().unique()
        if unique_wins.size == 1:
            return float(unique_wins[0])

    for category, df in categories.items():
        if category == "Skater":
            continue
        team_df = df[df[team_column] == team]
        if team_df.empty:
            continue
        wins_values = team_df[wins_column].dropna()
        if wins_values.empty:
            continue
        total_wins = wins_values.sum()
        if pd.notna(total_wins):
            return float(total_wins)

    return None


def calculate_wpa(data_file, coefficient_file, output_file):
    """Calculate Win Probability Added for skaters and goalies.

    The calculation produces per-player WPA values that sum to each team's wins
    across both skaters and goalies. Stat contributions are normalised by team
    totals, weighted by the provided coefficient tables, shifted to keep scores
    non-negative, and scaled so the combined WPA never exceeds the team's win
    total while maintaining a floor of zero.
    """

    data_tables = _load_category_tables(data_file)
    coeff_tables = _load_coefficients(coefficient_file)

    # Align available categories between data and coefficients
    categories = {
        category: _harmonize_columns(df.copy())
        for category, df in data_tables.items()
        if category in coeff_tables
    }

    if not categories:
        message = (
            "No overlapping categories found between data and " "coefficient files."
        )
        raise ValueError(message)

    team_column = "Team"
    wins_column = "W"

    for category, df in categories.items():
        if team_column not in df.columns:
            raise ValueError(
                "The data file is missing a '{column}' column for "
                "{category}s.".format(column=team_column, category=category)
            )
        if wins_column not in df.columns:
            raise ValueError(
                "The data file is missing a '{column}' column for "
                "{category}s.".format(column=wins_column, category=category)
            )

    coeff_dicts = {}
    tracked_stats = {}
    for category, df in categories.items():
        coeff_dict = _build_coefficient_dict(coeff_tables[category])
        coeff_dicts[category] = coeff_dict
        stats = [stat for stat in coeff_dict if stat in df.columns]
        tracked_stats[category] = stats
        if not stats:
            tracked_stats[category] = []

    wpa_series = {
        category: pd.Series(0.0, index=df.index) for category, df in categories.items()
    }

    # Build a master list of teams across all categories
    all_teams = set()
    for df in categories.values():
        all_teams.update(df[team_column].dropna().unique())

    team_total_wins = {}
    team_game_counts = {}

    for team in all_teams:
        team_wins = _determine_team_wins(
            team,
            categories,
            team_column,
            wins_column,
        )
        if team_wins is None:
            continue
        team_total_wins[team] = team_wins

        category_scores = {}
        category_totals = {}
        player_count = 0

        for category, df in categories.items():
            team_df = df[df[team_column] == team]
            if team_df.empty:
                continue

            if "GP" in team_df.columns:
                max_gp = team_df["GP"].dropna().max()
                if pd.notna(max_gp):
                    existing_games = team_game_counts.get(team, 0.0)
                    team_game_counts[team] = max(existing_games, float(max_gp))

            coeff_dict = coeff_dicts[category]
            stats = tracked_stats[category]

            if not stats:
                raw_scores = {idx: 1.0 for idx in team_df.index}
            else:
                team_totals = {stat: team_df[stat].fillna(0).sum() for stat in stats}
                raw_scores = {}
                for idx, row in team_df.iterrows():
                    score = 0.0
                    for stat in stats:
                        total = team_totals.get(stat, 0)
                        if not total:
                            continue
                        value = row.get(stat, 0)
                        if pd.isna(value):
                            continue
                        score += (value / total) * coeff_dict[stat]
                    raw_scores[idx] = score

                if category == "Skater":
                    raw_scores = _boost_point_contributions(
                        raw_scores,
                        team_df,
                    )

            adjusted_scores, total = _adjust_scores(raw_scores)
            category_scores[category] = adjusted_scores
            category_totals[category] = total
            player_count += len(adjusted_scores)

        if team_wins <= 0 or player_count == 0:
            # Either the team recorded no wins or we had no usable data; assign
            # zero WPA
            continue

        total_score_sum = sum(category_totals.values())
        num_categories = len(category_scores)

        category_weights = {}
        if total_score_sum <= 0:
            for category, scores in category_scores.items():
                category_weights[category] = (
                    len(scores) / player_count if player_count else 0.0
                )
        else:
            raw_weights = {}
            for category, scores in category_scores.items():
                score_share = (
                    category_totals[category] / total_score_sum
                    if total_score_sum
                    else 1.0 / num_categories
                )
                player_share = len(scores) / player_count if player_count else 0.0
                blended_weight = max(
                    CATEGORY_WEIGHT_BLEND * score_share
                    + (1.0 - CATEGORY_WEIGHT_BLEND) * player_share,
                    0.0,
                )
                participant_factor = (
                    len(scores) ** CATEGORY_PLAYER_POWER if len(scores) > 0 else 0.0
                )
                raw_weights[category] = blended_weight * participant_factor

            weight_sum = sum(raw_weights.values())
            if weight_sum <= 0:
                for category in raw_weights:
                    category_weights[category] = (
                        1.0 / num_categories if num_categories else 0.0
                    )
            else:
                adjusted_weights = {
                    category: (
                        (value / weight_sum) ** CATEGORY_WEIGHT_EXP
                        if weight_sum
                        else 0.0
                    )
                    for category, value in raw_weights.items()
                }
                adjusted_sum = sum(adjusted_weights.values())
                if adjusted_sum <= 0:
                    for category in raw_weights:
                        category_weights[category] = (
                            1.0 / num_categories if num_categories else 0.0
                        )
                else:
                    category_weights = {
                        category: value / adjusted_sum
                        for category, value in adjusted_weights.items()
                    }

        for category, scores in category_scores.items():
            category_total = category_totals.get(category, 0.0)
            category_weight = category_weights.get(category, 0.0)
            category_wins = team_wins * category_weight

            if category_total <= 0 or not scores:
                per_player = category_wins / len(scores) if len(scores) > 0 else 0.0
                for idx in scores:
                    wpa_series[category].at[idx] = per_player
                continue

            ratios = {
                idx: max(score / category_total, 0.0) ** PLAYER_RATIO_EXP
                for idx, score in scores.items()
            }
            ratio_sum = sum(ratios.values())
            if ratio_sum <= 0:
                per_player = category_wins / len(scores) if len(scores) > 0 else 0.0
                for idx in scores:
                    wpa_series[category].at[idx] = per_player
                continue

            for idx, ratio in ratios.items():
                wpa_series[category].at[idx] = category_wins * (ratio / ratio_sum)

        if 0.0 < TEAM_UNIFORM_BLEND <= 1.0 and player_count > 0:
            team_players = []
            for category, scores in category_scores.items():
                for idx in scores:
                    current = wpa_series[category].at[idx]
                    team_players.append((category, idx, current))

            if team_players:
                uniform_share = team_wins / len(team_players)
                for category, idx, current in team_players:
                    blended = (
                        1.0 - TEAM_UNIFORM_BLEND
                    ) * current + TEAM_UNIFORM_BLEND * uniform_share
                    wpa_series[category].at[idx] = blended

    team_win_rates = {}
    valid_rates = []
    for team, wins in team_total_wins.items():
        games = team_game_counts.get(team)
        if games and games > 0:
            rate = wins / games
            team_win_rates[team] = rate
            valid_rates.append(rate)

    if valid_rates:
        league_avg_rate = sum(valid_rates) / len(valid_rates)
    else:
        league_avg_rate = 1.0

    for team in team_total_wins:
        if team not in team_win_rates:
            team_win_rates[team] = league_avg_rate

    # Attach WPA values back onto their respective DataFrames
    for category, df in categories.items():
        df["WPA"] = wpa_series[category].fillna(0.0)
        if "GP" in df.columns:
            df = _move_column_after(df, "WPA", "GP")
        categories[category] = df.sort_values(
            by=[team_column, "WPA"], ascending=[True, False]
        )

    percent_lookup = _compute_team_win_percentages(
        categories,
        team_column,
        "WPA",
    )

    for category, df in categories.items():
        override_series = percent_lookup.get(category)
        updated_df = _attach_win_percentages(
            df,
            team_column,
            "WPA",
            override_percentages=override_series,
        )
        if category == "Skater":
            updated_df = _add_takeaway_difference(updated_df)
        categories[category] = updated_df

    # Prepare output directory
    output_dir = os.path.join("Results", "wpa")
    os.makedirs(output_dir, exist_ok=True)

    output_path = os.path.join(output_dir, output_file)

    combined_frames = [
        _build_combined_subset(df, category) for category, df in categories.items()
    ]

    with pd.ExcelWriter(output_path) as writer:
        for category, df in categories.items():
            sheet_name = f"{category}s"
            export_df = df.copy()

            if category == "Skater":
                export_df = export_df.drop(
                    columns=["Win % Value", "WPA"], errors="ignore"
                )

                column_names = list(export_df.columns)
                if "PIA" in column_names:
                    cutoff = column_names.index("PIA") + 1
                    base_columns = column_names[:cutoff]
                else:
                    base_columns = column_names

                if "Win %" in export_df.columns and "Win %" not in base_columns:
                    if "GP" in base_columns:
                        insert_at = base_columns.index("GP") + 1
                    else:
                        insert_at = len(base_columns)
                    base_columns.insert(insert_at, "Win %")

                export_columns = [
                    col for col in base_columns if col in export_df.columns
                ]
                if "Win %" in export_df.columns and "Win %" not in export_columns:
                    export_columns.append("Win %")

                export_df = export_df.loc[:, export_columns]
            elif category == "Goalie":
                export_df = export_df.drop(
                    columns=["Win % Value", "WPA"], errors="ignore"
                )

                column_names = list(export_df.columns)
                if "S3" in column_names:
                    cutoff = column_names.index("S3") + 1
                    base_columns = column_names[:cutoff]
                else:
                    base_columns = column_names

                if "Win %" in export_df.columns and "Win %" not in base_columns:
                    if "GP" in base_columns:
                        insert_at = base_columns.index("GP") + 1
                    else:
                        insert_at = len(base_columns)
                    base_columns.insert(insert_at, "Win %")

                export_columns = [
                    col for col in base_columns if col in export_df.columns
                ]
                if "Win %" in export_df.columns and "Win %" not in export_columns:
                    export_columns.append("Win %")

                export_df = export_df.loc[:, export_columns]
            else:
                export_df = export_df.copy()
                if "Win % Value" in export_df.columns:
                    export_df = export_df.drop(
                        columns=["Win % Value"],
                        errors="ignore",
                    )

            export_df.to_excel(writer, sheet_name=sheet_name, index=False)

        if len(combined_frames) > 1:
            combined_df = pd.concat(
                combined_frames,
                ignore_index=True,
                sort=False,
            )
            if "Win %" not in combined_df.columns:
                combined_df = _attach_win_percentages(
                    combined_df,
                    "Team Name",
                    "WPA",
                )
            else:
                combined_df["Win %"] = combined_df["Win %"].fillna(
                    combined_df["Win % Value"].apply(lambda value: f"{value:.2f}%")
                )

            roster_counts = combined_df.groupby("Team Name").size()
            if roster_counts.empty:
                avg_roster_size = 1.0
            else:
                avg_roster_size = float(roster_counts.mean())
                if avg_roster_size <= 0:
                    avg_roster_size = 1.0

            combined_df["Roster Size"] = (
                combined_df["Team Name"].map(roster_counts).fillna(avg_roster_size)
            )
            combined_df["Roster Size"] = combined_df["Roster Size"].replace(
                0, avg_roster_size
            )
            combined_df["Team Win % Value"] = (
                combined_df["Team Name"].map(team_win_rates).fillna(league_avg_rate)
            )
            combined_df["Team Win %"] = combined_df["Team Win % Value"].apply(
                lambda rate: f"{rate * 100:.2f}%"
            )

            combined_df = combined_df.sort_values(
                by=["Team Name", "Win % Value", "Player Name"],
                ascending=[True, False, True],
            ).reset_index(drop=True)

            combined_output = combined_df.copy()
            combined_output["Team Rank"] = (
                combined_output.groupby("Team Name", sort=False).cumcount() + 1
            )
            combined_output = combined_output[
                [
                    "Team Rank",
                    "Player Name",
                    "User",
                    "Team Name",
                    "POS",
                    "GP",
                    "Win %",
                ]
            ]

            combined_output.to_excel(
                writer,
                sheet_name="Combined",
                index=False,
            )

            position_sheets = {
                "Combined Defenders": {"D"},
                "Combined Forwards": {"C", "LW", "RW"},
                "Combined Goalies": {"G"},
            }

            for sheet_name, positions in position_sheets.items():
                ranking_df = _build_position_ranking(
                    combined_df, positions, avg_roster_size
                )
                if ranking_df.empty:
                    continue

                ranking_output = ranking_df.reset_index(drop=True)
                ranking_output.index = ranking_output.index + 1
                if "Win %" in ranking_output.columns:
                    ranking_output = ranking_output.drop(columns=["Win %"])
                ranking_output = ranking_output.rename(
                    columns={"Adjusted Win %": "Win %"}
                )
                ranking_output.insert(0, "Rank", ranking_output.index)
                ranking_output = ranking_output[
                    [
                        "Rank",
                        "Player Name",
                        "User",
                        "Team Name",
                        "POS",
                        "GP",
                        "Win %",
                    ]
                ]
                ranking_output.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False,
                )
        else:
            # Maintain backward compatibility with a single sheet workbook
            sole_category = next(iter(categories))
            categories[sole_category].to_excel(writer, index=False)

    _auto_fit_columns(output_path)

    print(f"WPA calculation complete. Results saved to {output_path}")


def calculate_wpa_batch(data_files, coefficient_file):
    """Process multiple data files with a shared coefficient workbook.

    Parameters
    ----------
    data_files : Iterable[str]
        One or more Excel files containing category statistics.
    coefficient_file : str
        Excel workbook containing the matching coefficient tables.

    Returns
    -------
    list[str]
        Absolute paths to the generated result workbooks.
    """

    outputs = []
    for data_file in data_files:
        if not data_file:
            continue
        output_name = os.path.basename(data_file)
        calculate_wpa(data_file, coefficient_file, output_name)
        outputs.append(os.path.join("Results", "wpa", output_name))
    return outputs


if __name__ == "__main__":
    # Create a file dialog for input selection
    root = tk.Tk()
    root.withdraw()

    print("Select up to four data files")
    data_files = filedialog.askopenfilenames(
        title="Select data files",
        filetypes=[
            ("Excel files", "*.xlsx *.xlsm *.xls"),
            ("All files", "*.*"),
        ],
    )

    data_files = list(data_files)
    if not data_files:
        print("No data files selected. Exiting.")
        raise SystemExit(0)

    if len(data_files) > 4:
        print(
            "More than four files selected; only the first four will be " "processed."
        )
        data_files = data_files[:4]

    print("Select the coefficient file")
    coefficient_file = filedialog.askopenfilename(title="Select the coefficient file")
    if not coefficient_file:
        print("No coefficient file selected. Exiting.")
        raise SystemExit(0)

    for data_file in data_files:
        if not data_file:
            continue
        output_name = os.path.basename(data_file)
        print(f"Processing {output_name}")
        calculate_wpa(data_file, coefficient_file, output_name)
