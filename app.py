"""
MLB Daily Betting Report - Streamlit Web App
Client opens URL in browser, enters odds, clicks button, downloads Excel report.
No Python installation needed on client side.
"""

import streamlit as st
import pandas as pd
import requests
import xlsxwriter
import io
import os
import re
from datetime import date, datetime, timedelta
from openpyxl import load_workbook

# ── Page config ───────────────────────────────────────────────────
st.set_page_config(
    page_title="MLB Daily Betting Report",
    page_icon="⚾",
    layout="wide"
)

# ── Import core logic from daily_report.py ────────────────────────
from daily_report import (
    load_historical_data, fetch_recent_results, compute_all_states,
    fetch_todays_schedule, get_team_state, build_game_row,
    check_scenarios, API_TO_CANONICAL, DIVISIONS, SCENARIO_DEFS,
    NEEDS_OPP_STREAK, NEEDS_OPP_ROAD_WP
)

# ── Cache heavy data loading ──────────────────────────────────────
@st.cache_data(ttl=3600, show_spinner="Loading historical data...")
def load_enriched_data(report_date_str):
    """Load and enrich all historical data. Cached for 1 hour."""
    df = load_historical_data()
    df = fetch_recent_results(df, date.fromisoformat(report_date_str))
    enriched = compute_all_states(df)
    return enriched

@st.cache_data(ttl=1800, show_spinner="Fetching today's schedule...")
def get_schedule(report_date_str):
    return fetch_todays_schedule(date.fromisoformat(report_date_str))

# ── Build Excel report in memory ──────────────────────────────────
def build_report_bytes(games, triggers, report_date, odds):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})

    fmt_title  = wb.add_format({'bold':True,'font_size':13,'font_color':'white',
                                 'bg_color':'#1F3864','align':'center','valign':'vcenter'})
    fmt_hdr    = wb.add_format({'bold':True,'font_color':'white','bg_color':'#2E75B6',
                                 'align':'center','border':1})
    fmt_away   = wb.add_format({'bg_color':'#E9EFF7','align':'left','border':1})
    fmt_home   = wb.add_format({'bg_color':'#FFFFFF','align':'left','border':1})
    fmt_away_c = wb.add_format({'bg_color':'#E9EFF7','align':'center','border':1})
    fmt_home_c = wb.add_format({'bg_color':'#FFFFFF','align':'center','border':1})
    fmt_fade   = wb.add_format({'bold':True,'bg_color':'#FFC7CE','font_color':'#9C0006',
                                 'align':'center','border':1})
    fmt_watch  = wb.add_format({'bold':True,'bg_color':'#FFEB9C','font_color':'#9C6500',
                                 'align':'center','border':1})
    fmt_none   = wb.add_format({'align':'center','border':1,'italic':True,'font_color':'#999999'})
    fmt_empty  = wb.add_format({'italic':True,'font_color':'#999999','align':'center'})

    def write_tab(ws, label, verdict_filter):
        ws.set_row(0, 22)
        ws.merge_range(0,0,0,4,
            f'MLB DAILY BETTING REPORT — {label} — {report_date.strftime("%A, %B %d, %Y")}',
            fmt_title)
        headers = ['Matchup','Team','Odds','Play','Scenario']
        widths  = [30,28,10,30,42]
        for ci,(h,w) in enumerate(zip(headers,widths)):
            ws.write(1,ci,h,fmt_hdr)
            ws.set_column(ci,ci,w)

        row = 2; written = 0
        for g in games:
            away, home = g['away_team'], g['home_team']
            matchup = f'{away} @ {home}'
            away_line = odds.get(away,'N/A')
            home_line = odds.get(home,'N/A')

            at = [t for t in triggers if t['team']==away and t['opponent']==home and t['verdict']==verdict_filter]
            ht = [t for t in triggers if t['team']==home and t['opponent']==away and t['verdict']==verdict_filter]
            if not at and not ht: continue
            written += 1

            def line_str(l):
                return f'+{l}' if isinstance(l,int) and l>0 else str(l)

            # Away row
            ap = ' | '.join(t['play'] for t in at) if at else '—'
            as_ = ' | '.join(f"#{t['scenario_id']} {t['scenario']}" for t in at) if at else '—'
            pf_a = fmt_fade if (at and at[0]['verdict']=='CLEAR FADE') else (fmt_watch if at else fmt_none)
            ws.write(row,0,f'↑ {matchup}',fmt_away)
            ws.write(row,1,f'(Away) {away}',fmt_away)
            ws.write(row,2,line_str(away_line),fmt_away_c)
            ws.write(row,3,ap, pf_a)
            ws.write(row,4,as_,fmt_away)
            row += 1

            # Home row
            hp = ' | '.join(t['play'] for t in ht) if ht else '—'
            hs = ' | '.join(f"#{t['scenario_id']} {t['scenario']}" for t in ht) if ht else '—'
            pf_h = fmt_fade if (ht and ht[0]['verdict']=='CLEAR FADE') else (fmt_watch if ht else fmt_none)
            ws.write(row,0,f'↓ {matchup}',fmt_home)
            ws.write(row,1,f'(Home) {home}',fmt_home)
            ws.write(row,2,line_str(home_line),fmt_home_c)
            ws.write(row,3,hp, pf_h)
            ws.write(row,4,hs,fmt_home)
            row += 2

        if written == 0:
            ws.merge_range(2,0,2,4,f'No {verdict_filter} scenarios triggered today.',fmt_empty)
        ws.freeze_panes(2,0)
        return written

    ws1 = wb.add_worksheet('Clear Fade');    ws1.set_tab_color('#FF0000')
    ws2 = wb.add_worksheet('Inconsistent'); ws2.set_tab_color('#FFCC00')
    n1 = write_tab(ws1,'CLEAR FADE','CLEAR FADE')
    n2 = write_tab(ws2,'INCONSISTENT','INCONSISTENT')

    # Summary
    ws3 = wb.add_worksheet('Summary'); ws3.set_tab_color('#2E75B6')
    ws3.set_row(0,22)
    ws3.merge_range(0,0,0,3,f'DAILY SUMMARY — {report_date.strftime("%A, %B %d, %Y")}',fmt_title)
    bf = wb.add_format({'bold':True,'border':1})
    cf = wb.add_format({'align':'center','border':1})
    rf = wb.add_format({'bold':True,'border':1,'bg_color':'#FFC7CE'})
    rc = wb.add_format({'align':'center','border':1,'bg_color':'#FFC7CE'})
    yf = wb.add_format({'bold':True,'border':1,'bg_color':'#FFEB9C'})
    yc = wb.add_format({'align':'center','border':1,'bg_color':'#FFEB9C'})
    ws3.write(1,0,'Total Games Today',bf);   ws3.write(1,1,len(games),cf)
    ws3.write(2,0,'Clear Fade Triggers',rf); ws3.write(2,1,sum(1 for t in triggers if t['verdict']=='CLEAR FADE'),rc)
    ws3.write(3,0,'Inconsistent Triggers',yf); ws3.write(3,1,sum(1 for t in triggers if t['verdict']=='INCONSISTENT'),yc)
    ws3.set_column(0,0,25); ws3.set_column(1,1,15)

    wb.close()
    output.seek(0)
    return output.getvalue(), n1, n2


# ── MAIN UI ───────────────────────────────────────────────────────

st.title("⚾ MLB Daily Betting Report")
st.markdown("---")

# Date selector
col1, col2 = st.columns([2,4])
with col1:
    report_date = st.date_input("Report Date", value=date.today())

report_date_str = report_date.isoformat()

# Load data
with st.spinner("Loading team data and schedule..."):
    try:
        enriched = load_enriched_data(report_date_str)
        games    = get_schedule(report_date_str)
    except Exception as e:
        st.error(f"Error loading data: {e}")
        st.stop()

if not games:
    st.warning("No games scheduled for this date.")
    st.stop()

st.success(f"✓ {len(games)} games scheduled for {report_date.strftime('%A, %B %d, %Y')}")
st.markdown("---")

# ── ODDS INPUT TABLE ──────────────────────────────────────────────
st.subheader("📋 Enter Today's Moneylines")
st.caption("Enter the moneyline for each team (e.g. +130 or -150). Leave blank if unavailable.")

# Build odds input form
odds_data = []
for g in games:
    odds_data.append({
        'Away Team': g['away_team'],
        'Home Team': g['home_team'],
        'Away Line': '',
        'Home Line': '',
    })

odds_df = pd.DataFrame(odds_data)
edited = st.data_editor(
    odds_df,
    column_config={
        'Away Team': st.column_config.TextColumn('Away Team', disabled=True, width='large'),
        'Home Team': st.column_config.TextColumn('Home Team', disabled=True, width='large'),
        'Away Line': st.column_config.NumberColumn('Away Line', help='e.g. 130 or -150', width='medium'),
        'Home Line': st.column_config.NumberColumn('Home Line', help='e.g. -130 or 150', width='medium'),
    },
    hide_index=True,
    use_container_width=True,
)

st.markdown("---")

# ── GENERATE REPORT ───────────────────────────────────────────────
if st.button("⚾ Generate Daily Report", type="primary", use_container_width=True):
    # Parse odds from editor
    odds = {}
    for _, row in edited.iterrows():
        if row['Away Line'] not in [None,'','nan']:
            try: odds[str(row['Away Team']).upper()] = int(row['Away Line'])
            except: pass
        if row['Home Line'] not in [None,'','nan']:
            try: odds[str(row['Home Team']).upper()] = int(row['Home Line'])
            except: pass

    if not odds:
        st.warning("Please enter at least some moneylines before generating the report.")
        st.stop()

    with st.spinner("Running scenario analysis..."):
        # Build opponent streak and road win% lookups
        opp_streaks = {}
        opp_road_wpct = {}
        for team in API_TO_CANONICAL.values():
            tdf = enriched[(enriched['team']==team) & (enriched['date'].dt.date < report_date)]
            if not tdf.empty:
                last = tdf.iloc[-1]
                opp_streaks[team] = last['streak_before'] + (1 if last['result']=='W' else -1)
                road = tdf[tdf['home_away']=='away']
                if not road.empty:
                    rw = (road['result']=='W').sum()
                    rl = (road['result']=='L').sum()
                    opp_road_wpct[team] = rw/(rw+rl) if (rw+rl)>0 else None

        # Check scenarios for each game
        all_triggers = []
        for g in games:
            away, home = g['away_team'], g['home_team']
            away_state = get_team_state(enriched, away, report_date)
            home_state = get_team_state(enriched, home, report_date)
            away_row = build_game_row(away_state, 'away', home, odds.get(away))
            home_row = build_game_row(home_state, 'home', away, odds.get(home))
            triggers = check_scenarios([away_row, home_row], opp_streaks, opp_road_wpct)
            all_triggers.extend(triggers)

        # Build Excel
        excel_bytes, n_fade, n_inc = build_report_bytes(games, all_triggers, report_date, odds)

    # Show summary
    st.markdown("---")
    st.subheader("📊 Results")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Games", len(games))
    m2.metric("🔴 Clear Fade Triggers", sum(1 for t in all_triggers if t['verdict']=='CLEAR FADE'))
    m3.metric("🟡 Inconsistent Triggers", sum(1 for t in all_triggers if t['verdict']=='INCONSISTENT'))

    # Show triggers in UI
    if all_triggers:
        st.markdown("### Clear Fade Plays")
        fade_list = [t for t in all_triggers if t['verdict']=='CLEAR FADE']
        if fade_list:
            fade_df = pd.DataFrame([{
                'Team': t['team'],
                'vs': t['opponent'],
                'H/A': t['home_away'].upper(),
                'Line': t['line'],
                'Play': t['play'],
                'Scenario': f"#{t['scenario_id']} {t['scenario']}",
            } for t in fade_list])
            st.dataframe(fade_df, use_container_width=True, hide_index=True)
        else:
            st.info("No Clear Fade scenarios triggered today.")

        st.markdown("### Inconsistent Plays")
        inc_list = [t for t in all_triggers if t['verdict']=='INCONSISTENT']
        if inc_list:
            inc_df = pd.DataFrame([{
                'Team': t['team'],
                'vs': t['opponent'],
                'H/A': t['home_away'].upper(),
                'Line': t['line'],
                'Play': t['play'],
                'Scenario': f"#{t['scenario_id']} {t['scenario']}",
            } for t in inc_list])
            st.dataframe(inc_df, use_container_width=True, hide_index=True)
        else:
            st.info("No Inconsistent scenarios triggered today.")

    # Download button
    st.markdown("---")
    fname = f'MLB_Daily_Report_{report_date.strftime("%Y-%m-%d")}.xlsx'
    st.download_button(
        label="📥 Download Excel Report",
        data=excel_bytes,
        file_name=fname,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
        type="primary",
    )
