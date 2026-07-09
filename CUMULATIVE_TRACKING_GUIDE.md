# 📊 Cumulative Scenario Tracking Guide

## Problem Solved

Previously, the **Scenario Performance** tab in each daily report would reset to zero because it only looked at that day's data. Now you can track cumulative results across the entire season!

## Solution Overview

We've implemented a **Master Results Database** system:
- A single `Master_Results.xlsx` file stores ALL results from the entire season
- Each daily report's **Scenario Performance** tab references this master file
- As you add results to the master file, all reports show updated cumulative statistics

---

## 🚀 Quick Start Guide

### First-Time Setup (One Time Only)

1. **Open the Streamlit App**
   - Run `streamlit run app.py` or access your deployed app

2. **Create Master Results File**
   - Scroll to bottom of the page
   - Click "🆕 Create Master Results File" button
   - The file `Master_Results.xlsx` will be created in your project folder

### Daily Workflow (Every Day)

#### Step 1: Generate Today's Report
1. Open the app and select today's date
2. Enter moneylines for all games
3. Click "⚾ Generate Daily Report"
4. Download the Excel report

#### Step 2: Copy Results to Master File
1. **Open your downloaded daily report**
   - File name: `MLB_Daily_Report_YYYY-MM-DD.xlsx`
   
2. **Go to "Results Tracker" tab**
   - This contains all of today's triggered scenarios
   
3. **Select and copy today's data**
   - Click on cell A4 (first data row)
   - Press `Ctrl+Shift+End` (Windows) or `Cmd+Shift+End` (Mac) to select all data
   - Copy: `Ctrl+C` (Windows) or `Cmd+C` (Mac)
   
4. **Open Master_Results.xlsx**
   - Located in same folder as the app
   
5. **Paste into Master file**
   - Click on the first empty row (Row 4 if first time, otherwise scroll down)
   - Paste: `Ctrl+V` (Windows) or `Cmd+V` (Mac)
   - **Save the file!** (`Ctrl+S` or `Cmd+S`)

#### Step 3: Track Game Results (Throughout the Day/Night)

1. **Open Master_Results.xlsx**
   
2. **Enter results as games complete**
   - In Column H ("Result"), enter:
     - `W` for wins
     - `L` for losses
   
3. **Save after each update**
   - Net P/L calculates automatically in Column I
   
4. **View Cumulative Stats**
   - Open ANY daily report (current or past)
   - Go to "Scenario Performance" tab
   - See updated season-long statistics!

---

## 📂 File Structure

```
your-project-folder/
├── app.py                          # Main Streamlit app
├── daily_report.py                 # Core analysis engine
├── master_results_manager.py       # NEW: Master file manager
├── Master_Results.xlsx             # NEW: Season-long results database
├── MLB_Daily_Report_2026-07-09.xlsx   # Today's report
├── MLB_Daily_Report_2026-07-10.xlsx   # Tomorrow's report
└── ...
```

---

## 📊 Understanding the Files

### Master_Results.xlsx
- **Purpose**: Stores ALL results from the entire season
- **Location**: Same folder as the app
- **Structure**: Same as "Results Tracker" tab in daily reports
- **Columns**:
  - Date, Team, H/A, Odds, Play, Scenario, Type
  - **Result**: You fill this in (W or L)
  - **Net P/L**: Auto-calculates based on result

### Daily Reports
- **Purpose**: Shows today's games and triggered scenarios
- **Tabs**:
  - 🟢 Clear Bet: High-confidence bet opportunities
  - 🔴 Clear Fade: High-confidence fade opportunities  
  - 🟡 Inconsistent: Scenarios with mixed historical performance
  - 📊 Summary: Overview of today's triggers
  - 📈 Results Tracker: Today's results (copy from here to Master file)
  - 📋 **Scenario Performance**: Cumulative stats (references Master file)

---

## 🎯 How Cumulative Tracking Works

### Before (Old System)
```
Daily Report #1 → Scenario Performance shows only Day 1 data
Daily Report #2 → Scenario Performance shows only Day 2 data
❌ No cumulative tracking!
```

### After (New System)
```
Day 1: Copy results to Master_Results.xlsx
Day 2: Copy results to Master_Results.xlsx  
Day 3: Copy results to Master_Results.xlsx

All Daily Reports → Scenario Performance tab → 
    References Master_Results.xlsx → 
    ✅ Shows cumulative season stats!
```

---

## 💡 Key Points

### ✅ DO:
- Keep `Master_Results.xlsx` in the same folder as your daily reports
- Copy results from EVERY daily report to the master file
- Save Master file after adding new results
- Back up your Master file regularly!
- You can enter W/L results at any time (doesn't have to be immediate)

### ❌ DON'T:
- Don't delete or rename `Master_Results.xlsx`
- Don't move the Master file to a different folder
- Don't skip copying a day's results (you'll lose tracking for that day)
- Don't edit the formulas in the Scenario Performance tab

---

## 📈 Viewing Your Stats

### In the Streamlit App
- Scroll to bottom after generating a report
- See "📊 Cumulative Season Tracking" section
- Expand "📈 View Master Results Summary" to see:
  - Total Wins/Losses
  - Win Percentage
  - Total Net P/L

### In Daily Reports (Excel)
- Open any daily report
- Go to "Scenario Performance" tab
- See cumulative statistics for all 36 scenarios:
  - Wins, Losses, Total plays
  - Win percentage
  - Net P/L
  - Season totals at the bottom

---

## 🔧 Troubleshooting

### "Scenario Performance tab shows all zeros"
**Cause**: Master_Results.xlsx is empty or doesn't have W/L results yet  
**Solution**: 
1. Open Master_Results.xlsx
2. Enter W or L in the "Result" column for completed games
3. Save the file
4. Reopen your daily report - stats will update!

### "Can't find Master_Results.xlsx"
**Cause**: File wasn't created or was moved/deleted  
**Solution**:
1. Open Streamlit app
2. Click "🆕 Create Master Results File" button
3. File will be created in the project folder

### "Formulas showing #REF! error"
**Cause**: Master_Results.xlsx is not in the same folder as daily report  
**Solution**:
1. Move Master_Results.xlsx back to project folder
2. Or regenerate the daily report

### "I forgot to copy a day's results"
**Solution**:
1. Open that day's daily report
2. Go to Results Tracker tab
3. Copy the data
4. Paste into Master_Results.xlsx
5. Enter the W/L results if known

---

## 🎓 Example Workflow

**Monday Morning (July 8, 2026)**
```
1. Open app, generate report for today
2. Download MLB_Daily_Report_2026-07-08.xlsx
3. Open report → Results Tracker tab
4. Copy rows 4-20 (today's 17 triggers)
5. Open Master_Results.xlsx
6. Paste at row 4 (first time) or next empty row
7. Save Master_Results.xlsx
```

**Monday Evening (Games Complete)**
```
1. Open Master_Results.xlsx
2. For each game, enter W or L in Result column
3. Watch Net P/L calculate automatically
4. Save file
```

**Tuesday Morning (View Stats)**
```
1. Generate new report for July 9
2. Go to Scenario Performance tab
3. See cumulative stats from both Monday AND Tuesday!
```

---

## 🆘 Support

If you encounter issues:
1. Check that Master_Results.xlsx is in the correct folder
2. Verify you've saved the Master file after adding results
3. Try regenerating the daily report
4. Check the app's "📖 How to Track Cumulative Results" section

---

## 🎉 Benefits

- ✅ **Season-Long Tracking**: See performance across all games, not just one day
- ✅ **Better Decisions**: Make informed bets based on cumulative scenario performance
- ✅ **Automatic Calculations**: P/L and win% calculate automatically
- ✅ **Simple Workflow**: Copy-paste daily results, that's it!
- ✅ **Historical Record**: Keep a complete record of all your betting activity

---

**Last Updated**: July 9, 2026  
**Version**: 2.0 (Master Results Database System)
