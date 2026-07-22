# Client Instructions — Cumulative Tracking

**Updated:** July 2026  
**Status:** Automated upload + master file workflow

---

## What Changed?

The **Scenario Performance** tab now shows **season-long cumulative stats** when you upload prior results before generating. The app builds and updates `Master_Results.xlsx` automatically — no manual copy/paste required.

---

## Daily Workflow (Streamlit App)

### 1. Open the app
```bash
streamlit run app.py
```

### 2. Upload prior results (optional from day 1, required from day 2+)
Upload either:
- **Master_Results.xlsx** from yesterday (recommended), or
- A saved **daily report** with W/L entered on the Results Tracker tab

Use **Clear cumulative data** to reset the session, or **Refresh game data cache** after updating the historical workbook or odds cache.

### 3. Generate today's report
1. Select today's date
2. Enter moneylines for scheduled games
3. Click **Generate Daily Report**

### 4. Download files
- **Daily report Excel** — today's triggers + Results Tracker (enter W/L here)
- **Master_Results.xlsx** — full season cumulative file (auto-merged)

### 5. Enter results
Open **Master_Results.xlsx** (or the daily report Results Tracker), enter **W** or **L** in column H after each game. Net P/L calculates automatically.

### 6. Next day
Upload the updated **Master_Results.xlsx** before generating again.

---

## Scenario Performance Tab

When master history is uploaded:
- **Season Cumulative** — prior days from master + today from Results Tracker
- Enter today's W/L on the Results Tracker tab; Scenario Performance updates via formulas

---

## CLI Alternative

```bash
python daily_report.py
```

Place `Master_Results.xlsx` in the project folder for CLI cumulative Scenario Performance. The CLI prints warnings if recent API results could not be fetched.

---

## Scatter Analysis Page

Requires `Martingale_Series_Analysis.xlsx` locally, or upload it in the app (Streamlit Cloud). Uses season win% **as of the selected date**.

---

## Tips

- Always download the updated **Master_Results.xlsx** after generating — it includes hidden columns for FADE P/L and opponent dedup
- Re-generating the same day **replaces** that day's triggers (does not duplicate)
- Legacy master files without hidden columns J/K still work; re-download from the app for full accuracy
