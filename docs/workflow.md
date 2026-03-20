# Hockey Stats Automation — Workflow

## Overview

This project is a multi-step hockey stats pipeline for VHL and VHLM regular season and playoff reporting.

The workflow:

1. Scrape game logs and build season-level results workbooks
2. Generate win probability coefficient data
3. Merge wins-added style metrics into the season results workbook
4. Generate team HTML pages and comparison HTML pages for publishing

The project is designed to turn raw game-log data into structured outputs with minimal manual work.

---

## Step 1 — Base season stat generation

Run one of the following depending on league and context:

- `vhl_stats_regular_season.py`
- `vhl_stats_playoffs.py`
- `vhlm_stats_regular_season.py`
- `vhlm_stats_playoffs.py`

These scripts:

- navigate the target hockey site
- crawl through game logs
- build season-level stats
- export the results to an Excel workbook

Typical output examples:

- `Results_VHL_100_Regular.xlsx`
- `Results_VHL_100_Playoffs.xlsx`
- `Results_VHLM_100_Regular.xlsx`

This is the base dataset used by downstream steps.

---

## Step 2 — Win probability coefficient generation

Run:

- `win_probability_added.py`

This script takes the season-level data and calculates win probability added factors used for later wins-added analysis.

Primary output:

- `Wins_Coefficient.xlsx`

This workbook acts as the intermediate coefficient dataset used in the next step.

---

## Step 3 — Wins-added analysis

Run:

- `analyze_wins.py`

This script takes:

- the base results workbook from Step 1
- the coefficient workbook from Step 2

It then calculates a wins-added style metric and writes the updated output to a downstream results file.

Typical behavior:

- reads a file like `Results_VHL_100_Regular.xlsx`
- adds a wins-added / WPA-related column
- saves the updated file to a `wpa` subdirectory

Example downstream output location:

- `Results/wpa/Results_VHL_100_Regular.xlsx`

---

## Step 4 — HTML publishing output

Run:

- `html_reports.py`

This script generates all HTML publishing output for the pipeline.

That includes:

- one HTML page per team
- position-based comparison pages

### Team report pages

The team report output displays all players for a team in HTML format, typically including:

- skater sections
- goalie sections
- wins-added / WPA-style output already merged into the dataset

Example output:

- `D.C._Dragons.html`

### Comparison pages

This same script also generates comparison pages for:

- defenders
- forwards
- goalies

These comparison outputs are designed to compare players against others in the same league and position group using the project’s wins-added / WPA-style metric.

---

## End-to-end summary

### Input
- hockey site game logs
- season context (league + regular season or playoffs)

### Intermediate outputs
- base Excel results workbook
- `Wins_Coefficient.xlsx`
- WPA-enhanced results workbook in a downstream `wpa` folder

### Final outputs
- team HTML report pages
- position-based HTML comparison pages

---

## Practical usage pattern

Typical run order:

1. run the appropriate season scraper
2. run `win_probability_added.py`
3. run `analyze_wins.py`
4. run `html_reports.py`

This produces a repeatable pipeline from scraped game logs to publishable HTML output.
