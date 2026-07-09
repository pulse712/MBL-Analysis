# 🚀 Quick Reference Card - Cumulative Tracking

**Print this out and keep it handy!**

---

## 📋 One-Time Setup (Do Once)

```
1. Open Streamlit app
2. Scroll to bottom
3. Click "Create Master Results File"
✅ Done!
```

---

## 📅 Daily Workflow (Every Day)

### Morning - Generate Report
```
☐ Open app
☐ Select today's date
☐ Enter moneylines
☐ Click "Generate Daily Report"
☐ Download Excel file
```

### Copy Today's Data (30 seconds)
```
☐ Open downloaded report
☐ Go to "Results Tracker" tab
☐ Select rows 4 and below (all data)
☐ Copy (Ctrl+C or Cmd+C)
☐ Open Master_Results.xlsx
☐ Find first empty row
☐ Paste (Ctrl+V or Cmd+V)
☐ Save file (Ctrl+S or Cmd+S)
```

### Track Results (As Games Complete)
```
☐ Open Master_Results.xlsx
☐ In Column H, enter:
   • W for wins
   • L for losses
☐ Save file
☐ Net P/L auto-calculates!
```

### View Your Stats
```
☐ Open ANY daily report
☐ Go to "Scenario Performance" tab
☐ See cumulative season stats! 🎉
```

---

## 🎯 Key Points

### ✅ ALWAYS Do This:
- Copy results to Master file EVERY day
- Save Master file after entering W/L
- Keep Master file in same folder as app
- Back up Master file weekly

### ❌ NEVER Do This:
- Delete or rename Master_Results.xlsx
- Move Master file to different folder
- Skip a day's copy-paste
- Edit formulas in Scenario Performance tab

---

## 📂 File Locations

**Master File**: Same folder as app.py  
**Daily Reports**: Downloads folder  
**Path shown in**: Bottom of Streamlit app

---

## ❓ Quick Troubleshooting

### "Scenario Performance shows zeros"
→ Master file is empty or no W/L entered yet  
→ Add W/L results and save

### "Can't find Master_Results.xlsx"
→ Click "Create Master Results File" in app  
→ File created in project folder

### "#REF! error in formulas"
→ Master file not in correct folder  
→ Move Master_Results.xlsx to app folder

---

## 📊 What You'll See

### In Master_Results.xlsx:
```
Row 1-2: Title and instructions
Row 3:   Column headers
Row 4+:  Your betting results
         (Copy daily data here!)
```

### In Scenario Performance Tab:
```
Column D: Wins
Column E: Losses  
Column F: Total plays
Column G: Win %
Column H: Net P/L
Row 41:  SEASON TOTALS
```

---

## 🎯 Example Day

**9:00 AM** - Generate report for today  
**9:02 AM** - Copy Results Tracker → Master file  
**2:00 PM** - Enter W/L for afternoon games  
**7:00 PM** - Enter W/L for evening games  
**9:00 PM** - View stats in Scenario Performance tab

**Total Time**: < 5 minutes per day

---

## 📱 Quick Commands

| Action | Windows | Mac |
|--------|---------|-----|
| Copy | Ctrl + C | Cmd + C |
| Paste | Ctrl + V | Cmd + V |
| Save | Ctrl + S | Cmd + S |
| Select All Below | Ctrl + Shift + End | Cmd + Shift + End |

---

## 🆘 Need Help?

1. Check in-app instructions (bottom of Streamlit page)
2. Read `CLIENT_INSTRUCTIONS.md`
3. See `CUMULATIVE_TRACKING_GUIDE.md` for details

---

## 🎉 Success Checklist

After your first week, you should have:

- [ ] Master_Results.xlsx with 5-7 days of data
- [ ] W/L results entered for completed games
- [ ] Cumulative stats showing in any report
- [ ] Running total of Net P/L
- [ ] Clear trend of scenario performance

---

**Remember**: The more data you track, the more valuable your cumulative statistics become!

---

*Keep this reference card handy for your daily workflow!*

**Version 2.0** | Last Updated: July 9, 2026
