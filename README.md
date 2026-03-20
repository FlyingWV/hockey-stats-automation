# Hockey Stats Automation

Python automation project for scraping hockey game logs, generating advanced player metrics, and publishing team-level and comparison HTML reports.

## Overview

This project automates a multi-step hockey stats workflow for VHL and VHLM regular season and playoff data.

The pipeline:

1. Scrapes game log data for a selected league/season context
2. Builds season-level Excel results workbooks
3. Calculates win-probability-based coefficient data
4. Adds wins-added style metrics to the results dataset
5. Generates team HTML pages and position-based comparison reports for publishing

The workflow turns raw game-log data into structured season outputs and publishable reporting pages with minimal manual work.

### Project Structure

src/ -- scripts
samples/ -- curated sample outputs
docs/ --workflow and output notes
output/ -- generated output folder

## Setup

### Requirements
- Python 3.11+ recommended

### Install dependencies

```bash
pip install -r requirements.txt

## Main Scripts

### Base stat generation
Run one of the following depending on league and context:

- `vhl_stats_regular_season.py`
- `vhl_stats_playoffs.py`
- `vhlm_stats_regular_season.py`
- `vhlm_stats_playoffs.py`

These scripts crawl game logs, build season-level skater and goalie statistics, and export a base results workbook.

### WPA coefficient generation
- `win_probability_added.py`

Builds the intermediate `Wins_Coefficient.xlsx` dataset used for downstream wins-added calculations.

### Wins-added analysis
- `analyze_wins.py`

Takes the base season results workbook plus the coefficient workbook and writes an enhanced results output with wins-added style metrics included.

### HTML report generation
- `html_reports.py`

Generates:
- team-level HTML report pages
- comparison HTML pages for defenders, forwards, and goalies

## Output Structure

### Excel outputs
The pipeline creates season-level Excel workbooks such as:
- `Results_VHL_100_Regular.xlsx`
- `Results_VHLM_100_Regular.xlsx`
- `Wins_Coefficient.xlsx`

A WPA-enhanced version of the results workbook is also produced in a downstream output step.

### HTML outputs
The reporting step creates:
- team HTML pages for individual clubs
- comparison HTML pages for position-based analysis

Comparison reports normalize player impact within the league and position grouping being compared.

## Sample Outputs

This repo includes sample outputs demonstrating:
- season results workbook
- wins coefficient workbook
- team HTML report page
- defender / forward / goalie comparison pages

## Tech Used

- Python
- web scraping / HTTP requests
- HTML parsing
- Excel output generation
- metric calculation workflows
- HTML report generation

## Why this project exists

This project demonstrates a repeatable sports data pipeline that combines:

- structured web scraping
- season-level stat aggregation
- custom metric generation
- Excel-based output workflows
- HTML publishing output

## Notes

This is a personal hockey stats automation project built to support recurring reporting and publishing workflows.
