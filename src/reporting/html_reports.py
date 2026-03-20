import pandas as pd
import re
from pathlib import Path
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from zipfile import BadZipFile

SKATER_SHEET = "Skaters"
GOALIE_SHEET = "Goalies"
TEAM_COMBINED_SHEET = "Combined"
LEAGUE_WIDE_SHEETS = {
    "Combined Defenders": "Combined Defenders",
    "Combined Forwards": "Combined Forwards",
    "Combined Goalies": "Combined Goalies",
}
SKATER_MAX_COLUMN = "PIA"
GOALIE_MAX_COLUMN = "S3"
PLAYER_LINK_TEMPLATE = (
    "https://vhlportal.com/players/playerfocus/{player_number}"
)
USER_LINK_TEMPLATE = "https://vhlportal.com/user/{user_id}"


def read_excel_file(file_path):
    combined_tables = {}
    team_combined = pd.DataFrame()
    try:
        with pd.ExcelFile(file_path, engine="openpyxl") as workbook:
            skaters = workbook.parse(SKATER_SHEET)
            goalies = workbook.parse(GOALIE_SHEET)
            sheet_names = set(workbook.sheet_names)

            if TEAM_COMBINED_SHEET in sheet_names:
                team_combined = workbook.parse(TEAM_COMBINED_SHEET)

            for label, sheet_name in LEAGUE_WIDE_SHEETS.items():
                if sheet_name not in sheet_names:
                    combined_tables[label] = pd.DataFrame()
                    continue
                combined_tables[label] = workbook.parse(sheet_name)
    except BadZipFile as exc:
        raise ValueError(
            f"The file '{file_path}' is not a valid Excel workbook."
        ) from exc
    except ValueError as exc:
        raise ValueError(
            f"One of the required sheets is missing in '{file_path}': {exc}"
        ) from exc
    except Exception as exc:
        raise ValueError(
            f"An error occurred while reading '{file_path}': {exc}"
        ) from exc

    skaters.columns = skaters.columns.str.strip()
    goalies.columns = goalies.columns.str.strip()
    if SKATER_MAX_COLUMN in skaters.columns:
        skaters = skaters.loc[:, :SKATER_MAX_COLUMN]
    if GOALIE_MAX_COLUMN in goalies.columns:
        goalies = goalies.loc[:, :GOALIE_MAX_COLUMN]

    team_combined = team_combined.copy()
    if not team_combined.empty:
        team_combined.columns = team_combined.columns.str.strip()

    cleaned_combined = {}
    for label, df in combined_tables.items():
        copy_df = df.copy()
        copy_df.columns = copy_df.columns.str.strip()
        cleaned_combined[label] = copy_df

    return {
        "Skaters": skaters,
        "Goalies": goalies,
        "Combined": team_combined,
        "league_wide": cleaned_combined,
    }


def map_player_data(player_data_file):
    try:
        player_data = pd.read_csv(player_data_file, low_memory=False)
    except Exception as exc:
        raise ValueError(
            f"An error occurred while reading '{player_data_file}': {exc}"
        ) from exc

    player_data.columns = player_data.columns.str.strip()
    required = {"Player Name", "Player Number", "User", "User ID"}
    missing = required.difference(player_data.columns)
    if missing:
        missing_list = ", ".join(sorted(missing))
        raise ValueError(
            "The Player Data file is missing required columns: "
            f"{missing_list}."
        )

    player_data["Player Name"] = (
        player_data["Player Name"].astype(str).str.strip()
    )
    player_data["User"] = player_data["User"].astype(str).str.strip()

    def normalize_identifier(value):
        if pd.isna(value):
            return ""
        text = str(value).strip()
        if not text:
            return ""
        try:
            return str(int(float(text)))
        except ValueError:
            return text

    player_data["User ID"] = player_data["User ID"].apply(normalize_identifier)
    player_data["Player Number"] = player_data["Player Number"].apply(
        normalize_identifier
    )

    player_numbers = dict(
        zip(player_data["Player Name"], player_data["Player Number"])
    )
    player_to_user = dict(
        zip(
            player_data["Player Name"],
            zip(player_data["User"], player_data["User ID"]),
        )
    )
    return player_numbers, player_to_user


def _player_link(name, player_map):
    number = player_map.get(name, "")
    if not number:
        return name
    url = PLAYER_LINK_TEMPLATE.format(player_number=number)
    return f'<a href="{url}">{name}</a>'


def _user_link(
    player_name,
    player_to_user_map,
    fallback,
):
    user_info = player_to_user_map.get(player_name)
    if not user_info:
        if fallback.strip():
            return fallback
        return ""
    user_name, user_id = user_info
    if user_id and user_name:
        url = USER_LINK_TEMPLATE.format(user_id=user_id)
        return f'<a href="{url}">{user_name}</a>'
    return user_name or fallback


def generate_html_table(
    data,
    player_map,
    player_to_user_map,
    title,
    name_col,
):
    if name_col not in data.columns:
        raise ValueError(
            f"The data for '{title}' is missing the '{name_col}' column."
        )

    table_df = data.copy().reset_index(drop=True)
    original_names = table_df[name_col].astype(str).str.strip()
    original_list = original_names.tolist()
    table_df[name_col] = [
        _player_link(name, player_map) for name in original_list
    ]

    user_links = []
    for idx, row in enumerate(table_df.itertuples(index=False)):
        fallback_value = getattr(row, "User", "")
        if fallback_value is None:
            fallback_text = ""
        else:
            fallback_text = str(fallback_value)
        linked_user = _user_link(
            original_list[idx],
            player_to_user_map,
            fallback_text,
        )
        user_links.append(str(linked_user))
    table_df["User"] = user_links

    html_table = table_df.to_html(index=False, escape=False)
    return f"<h2>{title}</h2>" + html_table


def select_file(prompt):
    Tk().withdraw()
    print(prompt)
    return askopenfilename(title=prompt)


def extract_season_from_filename(filename):
    """Extract season information from a results filename such as
    'Results_VHL_100_Regular.xlsx'.
    """
    match = re.search(r'_(\d+)_', filename)
    if match:
        return f"S{match.group(1)}"
    return "Unknown_Season"


def _extract_team_names(dataset):
    if not dataset:
        return set()

    skaters = dataset.get("Skaters", pd.DataFrame())
    if skaters.empty or "Team Name" not in skaters.columns:
        return set()

    teams = skaters["Team Name"].dropna().astype(str).str.strip().unique()
    return {team for team in teams if team}


def build_league_sections(
    team,
    league_tag,
    all_data,
    player_map,
    player_to_user_map,
):
    sections = []
    for title, dataset in all_data.items():
        if league_tag not in title:
            continue

        skaters = dataset["Skaters"]
        goalies = dataset["Goalies"]
        team_skaters = skaters[skaters["Team Name"] == team]
        if not team_skaters.empty:
            sections.append(
                generate_html_table(
                    team_skaters,
                    player_map,
                    player_to_user_map,
                    f"{title} - Skaters",
                    "Player Name",
                )
            )

        team_goalies = goalies[goalies["Team Name"] == team]
        if not team_goalies.empty:
            sections.append(
                generate_html_table(
                    team_goalies,
                    player_map,
                    player_to_user_map,
                    f"{title} - Goalies",
                    "Goalie Name",
                )
            )

        combined_df = dataset.get("Combined", pd.DataFrame())
        if not combined_df.empty:
            team_combined = combined_df[combined_df["Team Name"] == team]
            if not team_combined.empty:
                sections.append(
                    generate_html_table(
                        team_combined,
                        player_map,
                        player_to_user_map,
                        f"{title} - Combined (Team Win %)",
                        "Player Name",
                    )
                )

    return sections


def generate_league_wide_reports(
    all_data,
    player_map,
    player_to_user_map,
    base_output_dir,
    season,
):
    for dataset_name, dataset in all_data.items():
        league_tables = dataset.get("league_wide", {})
        for label, table_df in league_tables.items():
            if table_df.empty:
                continue

            title = f"{dataset_name} - {label}"
            if "Player Name" in table_df.columns:
                name_column = "Player Name"
            elif "Goalie Name" in table_df.columns:
                name_column = "Goalie Name"
            else:
                raise ValueError(
                    "Unable to determine the player name column for "
                    f"'{dataset_name}' '{label}'."
                )

            section_html = generate_html_table(
                table_df,
                player_map,
                player_to_user_map,
                title,
                name_column,
            )

            output_html = f"<html><body>{section_html}</body></html>"
            
            # Extract league and season type from dataset name
            if "VHLM" in dataset_name:
                league = "VHLM"
            elif "VHL" in dataset_name:
                league = "VHL"
            else:
                league = "Unknown"
            if "Regular" in dataset_name:
                season_type = "Regular"
            else:
                season_type = "Playoffs"

            comparisons_dir = (
                base_output_dir
                / "comparisons"
                / season
                / league
                / season_type
            )
            comparisons_dir.mkdir(parents=True, exist_ok=True)
            
            # Determine file name based on label (simplified naming)
            if "Combined Defenders" in label:
                filename = "defenders.html"
            elif "Combined Forwards" in label:
                filename = "forwards.html"
            elif "Combined Goalies" in label:
                filename = "goalies.html"
            else:
                # Fallback for any other combined reports
                filename = f"{label.replace(' ', '_').lower()}.html"
            
            output_file = comparisons_dir / filename
            output_file.write_text(output_html, encoding="utf-8")
            print(
                "Generated league-wide HTML for",
                f"{dataset_name} {label}:",
                output_file.resolve(),
            )


def main():
    vhl_regular_file = select_file("Select the VHL Regular Season file")
    vhl_playoffs_file = select_file("Select the VHL Playoffs file")
    vhlm_regular_file = select_file("Select the VHLM Regular Season file")
    vhlm_playoffs_file = select_file("Select the VHLM Playoffs file")
    player_data_file = select_file("Select the Player Data file")

    if not vhl_regular_file:
        print("A VHL Regular Season file is required. Exiting.")
        return

    if not player_data_file:
        print("A Player Data file is required. Exiting.")
        return

    player_map, player_to_user_map = map_player_data(player_data_file)

    # Extract season from the first available file
    season_source = next(
        (
            file_path
            for file_path in (
                vhl_regular_file,
                vhlm_regular_file,
                vhl_playoffs_file,
                vhlm_playoffs_file,
            )
            if file_path
        ),
        "",
    )
    season = extract_season_from_filename(Path(season_source).name)
    
    base_output_dir = Path("html")
    base_output_dir.mkdir(exist_ok=True)

    all_data = {}

    all_data["VHL Regular Season"] = read_excel_file(vhl_regular_file)

    if vhl_playoffs_file:
        all_data["VHL Playoffs"] = read_excel_file(vhl_playoffs_file)
    else:
        print("No VHL Playoffs file selected; skipping VHL playoff reports.")

    if vhlm_regular_file:
        all_data["VHLM Regular Season"] = read_excel_file(vhlm_regular_file)
    else:
        print("No VHLM Regular Season file selected; skipping VHLM regular reports.")

    if vhlm_playoffs_file:
        all_data["VHLM Playoffs"] = read_excel_file(vhlm_playoffs_file)
    elif vhlm_regular_file:
        print("No VHLM Playoffs file selected; skipping VHLM playoff reports.")

    vhlm_teams = _extract_team_names(all_data.get("VHLM Regular Season"))
    vhl_teams = _extract_team_names(all_data.get("VHL Regular Season"))

    # Generate team reports with new directory structure
    all_teams = sorted(vhlm_teams.union(vhl_teams))
    for team in all_teams:
        # Determine which leagues this team belongs to
        team_leagues = []
        if team in vhlm_teams:
            team_leagues.append("VHLM")
        if team in vhl_teams:
            team_leagues.append("VHL")
        
        # Generate reports for each league the team is in
        for league in team_leagues:
            # Regular Season
            regular_sections = build_league_sections(
                team,
                league,
                {k: v for k, v in all_data.items() if league in k and "Regular" in k},
                player_map,
                player_to_user_map,
            )
            
            if regular_sections:
                output_html = "<html><body>" + "".join(regular_sections) + "</body></html>"
                team_dir = base_output_dir / season / league / "Regular"
                team_dir.mkdir(parents=True, exist_ok=True)
                output_file = team_dir / f"{team.replace(' ', '_')}.html"
                output_file.write_text(output_html, encoding="utf-8")
                print(
                    f"Generated Regular Season HTML for {league} team",
                    team,
                    "at",
                    output_file.resolve(),
                )
            
            # Playoffs
            playoff_sections = build_league_sections(
                team,
                league,
                {k: v for k, v in all_data.items() if league in k and "Playoffs" in k},
                player_map,
                player_to_user_map,
            )
            
            if playoff_sections:
                output_html = "<html><body>" + "".join(playoff_sections) + "</body></html>"
                team_dir = base_output_dir / season / league / "Playoffs"
                team_dir.mkdir(parents=True, exist_ok=True)
                output_file = team_dir / f"{team.replace(' ', '_')}.html"
                output_file.write_text(output_html, encoding="utf-8")
                print(
                    f"Generated Playoffs HTML for {league} team",
                    team,
                    "at",
                    output_file.resolve(),
                )

    generate_league_wide_reports(
        all_data,
        player_map,
        player_to_user_map,
        base_output_dir,
        season,
    )


if __name__ == "__main__":
    main()
