# MLB Daily Betting Report

A Streamlit web app that generates a daily MLB betting report based on 36 scenario filters.

## How it works

1. Opens today's MLB schedule automatically
2. Client enters today's moneylines
3. App checks each game against 36 betting scenarios
4. Generates a downloadable Excel report with Clear Fade and Inconsistent plays

## Files

- `app.py` — Streamlit web interface
- `daily_report.py` — Core analysis engine (data loading, state computation, scenario filters)
- `MLB Data 2023-2026.xlsx` — Historical game data (2023-2026 seasons)
- `requirements.txt` — Python dependencies

## Deploy on Streamlit Cloud

1. Fork this repo
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set main file to `app.py`
5. Deploy
