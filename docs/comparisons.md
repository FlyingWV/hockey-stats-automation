# Hockey Stats Automation — Comparison Reports

## Overview

The project includes position-based comparison reports generated as HTML pages.

These comparison outputs are created by:

- `html_reports.py`

There is no separate comparison-only script in the current workflow. The main HTML publishing script creates both:

- team report pages
- comparison report pages

---

## Purpose of the comparison reports

The comparison pages are designed to compare players against one another within a shared context.

The comparison structure groups players by:

- league
- season context
- position category

The position categories currently described are:

- defenders
- forwards
- goalies

These pages help answer questions like:

- how does one player compare to others at the same position?
- where does a player rank within the league context being analyzed?
- how does a player's wins-added / WPA-style value compare to peers?

---

## Metric concept

The comparison reports are based on the project’s wins-added / WPA-style analysis.

The workflow is:

1. scrape base season results
2. generate coefficient data
3. write wins-added style output into the enhanced workbook
4. generate comparison HTML pages from that enhanced dataset

The comparison pages are intended to compare players using a wins-added style metric normalized to the league and position grouping being analyzed.

That makes them more useful than raw totals alone.

---

## Folder flow

The working project structure includes comparison output paths such as:

- `html/comparisons/S100/VHL/Regular/`
- `html/comparisons/S100/VHL/Playoffs/`
- `html/comparisons/S100/VHLM/Regular/`
- `html/comparisons/S100/VHLM/Playoffs/`

Within those folders, the generated files include:

- `defenders.html`
- `forwards.html`
- `goalies.html`

This makes the comparison reports easy to separate by:

- season
- league
- regular season vs playoffs
- position group

---

## Why the comparison reports matter

The comparison outputs make the project more than a scraping tool.

They show that the pipeline supports:

- downstream metric interpretation
- normalized player comparison
- league/position-specific reporting
- publishable HTML presentation

This is one of the stronger features of the project because it demonstrates a full pipeline from raw site data to comparative published analysis.

---
