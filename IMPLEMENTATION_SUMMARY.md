# Implementation Summary

## Architecture (July 2026)

| Component | Location |
|-----------|----------|
| Scenario engine + CLI | `daily_report.py` |
| Streamlit Excel builder | `report_builder.py` |
| Master file I/O | `master_results_manager.py` |
| Streamlit UI | `pages/1_Daily_Report.py` |
| Scatter analysis | `pages/2_Scatter_Analysis.py` |

## Key Features

- Automated master merge on generate (upload → generate → download)
- Hidden columns J (PayoutLine) and K (Opponent) for accurate FADE P/L and DH dedup
- API backfill with warnings; doubleheader both games kept
- Session controls: clear cumulative data, refresh cache
- 28 automated tests in `test_master_results.py`

See `CLIENT_INSTRUCTIONS.md` for the current daily workflow.
