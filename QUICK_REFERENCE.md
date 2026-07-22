# Quick Reference

## Streamlit
```bash
streamlit run app.py
```

## CLI
```bash
python daily_report.py
```

## Tests
```bash
python test_master_results.py
```

## Daily Steps
1. Upload `Master_Results.xlsx` (or yesterday's saved report)
2. Enter moneylines → Generate
3. Download daily report + updated master
4. Enter W/L in master file after games

## Key Modules
| Module | Role |
|--------|------|
| `daily_report.py` | Scenario engine, CLI report, API backfill |
| `report_builder.py` | Streamlit Excel download (in-memory) |
| `master_results_manager.py` | Master file parse/append/stats |
| `pages/1_Daily_Report.py` | Streamlit UI |
| `pages/2_Scatter_Analysis.py` | Dog vs fav scatter (needs Martingale file) |

## Troubleshooting
- **Stale data** → Refresh game data cache
- **Wrong cumulative** → Clear cumulative data, re-upload master
- **API gaps** → Yellow warning banner; check network and retry
- **No history for team** → Warning on generate; check `MLB Data 2023-2026.xlsx`
