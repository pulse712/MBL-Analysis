# MLB Daily Betting Report

A Streamlit web app that generates a daily MLB betting report based on 36 scenario filters with **season-long cumulative tracking**.

## How it works

1. Opens today's MLB schedule automatically
2. Client enters today's moneylines
3. App checks each game against 36 betting scenarios
4. Generates a downloadable Excel report with Clear Bet, Clear Fade, and Inconsistent plays
5. **NEW**: Tracks cumulative scenario performance across the entire season

## Key Features

### 🎯 Daily Analysis
- Automatically fetches today's MLB schedule
- Analyzes each game against 36 proven betting scenarios
- Generates color-coded recommendations:
  - 🟢 **Clear Bet**: High-confidence betting opportunities
  - 🔴 **Clear Fade**: High-confidence fade opportunities
  - 🟡 **Inconsistent**: Scenarios with mixed historical performance

### 📊 Season-Long Tracking (NEW!)
- **Master Results Database**: Single file tracks ALL results across the season
- **Cumulative Statistics**: See scenario performance for the entire season, not just one day
- **Automatic Calculations**: Win%, Net P/L, and totals update automatically
- **Simple Workflow**: Copy daily results to master file, that's it!

## Files

- `app.py` — Streamlit web interface
- `daily_report.py` — Core analysis engine (data loading, state computation, scenario filters)
- `master_results_manager.py` — **NEW**: Manages cumulative season tracking
- `Master_Results.xlsx` — **NEW**: Season-long results database (created automatically)
- `MLB Data 2023-2026.xlsx` — Historical game data (2023-2026 seasons)
- `requirements.txt` — Python dependencies
- `CUMULATIVE_TRACKING_GUIDE.md` — **NEW**: Complete guide for season tracking

## Quick Start

### Installation
```bash
pip install -r requirements.txt
```

### Run the App
```bash
streamlit run app.py
```

### First-Time Setup
1. Open the app in your browser
2. Scroll to bottom and click "🆕 Create Master Results File"
3. This creates `Master_Results.xlsx` for season-long tracking

### Daily Workflow
1. **Generate Report**: Select date, enter moneylines, generate report
2. **Copy to Master**: Copy "Results Tracker" data to `Master_Results.xlsx`
3. **Track Results**: Enter W/L as games complete
4. **View Stats**: See cumulative performance in "Scenario Performance" tab

See `CUMULATIVE_TRACKING_GUIDE.md` for detailed instructions.

## Report Structure

Each daily report contains 6 tabs:

1. **🟢 Clear Bet** - High-confidence bet opportunities
2. **🔴 Clear Fade** - High-confidence fade opportunities
3. **🟡 Inconsistent** - Mixed-performance scenarios
4. **📊 Summary** - Overview with top plays
5. **📈 Results Tracker** - Track today's results (copy to Master file)
6. **📋 Scenario Performance** - Cumulative season statistics

## Cumulative Tracking System

### How It Works
```
Daily Report → Results Tracker → Copy to Master_Results.xlsx → 
Scenario Performance tab reads from Master file → Shows cumulative stats!
```

### Benefits
- ✅ Track performance across the entire season
- ✅ Make data-driven decisions based on cumulative results
- ✅ Automatic Win%, Net P/L calculations
- ✅ Complete historical record

## Deploy on Streamlit Cloud

1. Fork this repo
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set main file to `app.py`
5. Deploy

**Note**: After deployment, create Master_Results.xlsx file using the app interface.

## Support

For questions about cumulative tracking, see `CUMULATIVE_TRACKING_GUIDE.md`.

## What's New in v2.0

- 🎉 **Master Results Database** - Season-long cumulative tracking
- 📊 **Enhanced Scenario Performance** - References external master file
- 📖 **In-App Instructions** - Step-by-step guidance for tracking results
- 📈 **Statistics Dashboard** - View cumulative stats in the app

---

**Version**: 2.0  
**Last Updated**: July 9, 2026
