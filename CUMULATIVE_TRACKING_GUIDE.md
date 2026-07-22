# Cumulative Tracking Guide

## Overview

Season results live in **Master_Results.xlsx**. The Streamlit app merges uploaded history with each new day's triggers automatically.

## Files

| File | Purpose |
|------|---------|
| `Master_Results.xlsx` | Season cumulative W/L (hidden cols J=PayoutLine, K=Opponent) |
| `MLB_Daily_Report_YYYY-MM-DD.xlsx` | Single-day report with Results Tracker |
| `MLB Data 2023-2026.xlsx` | Historical game logs (auto-updated from API) |

## Upload Options

1. **Master_Results.xlsx** — preferred after day 1
2. **Saved daily report** — Results Tracker tab with W/L filled in

## Session Controls

- **Clear cumulative data** — removes uploaded master from this browser session
- **Refresh game data cache** — reloads historical data + API backfill (use after editing source files)

## Dedup Rules

Rows are unique by: `(date, team, scenario, opponent, payout line)`. Same-day re-generate replaces prior rows for that date.

## Legacy Files

Older masters without columns J/K trigger a warning on upload. FADE P/L and doubleheader dedup work best after re-downloading from the app.
