"""
MLB Daily Betting Report - Streamlit Web App
"""
import streamlit as st
import pandas as pd
import xlsxwriter
import io
from datetime import date, datetime
from daily_report import (
    load_historical_data, fetch_recent_results, compute_all_states,
    fetch_todays_schedule, get_team_state, build_game_row,
    check_scenarios, API_TO_CANONICAL, SCENARIO_DEFS,
    NEEDS_OPP_STREAK, NEEDS_OPP_ROAD_WP, title_case, fmt_line
)

st.set_page_config(page_title="MLB Daily Betting Report", page_icon="⚾", layout="wide")

@st.cache_data(ttl=3600, show_spinner="Loading historical data...")
def load_enriched_data(report_date_str):
    df = load_historical_data()
    df = fetch_recent_results(df, date.fromisoformat(report_date_str))
    return compute_all_states(df)

@st.cache_data(ttl=1800, show_spinner="Fetching today's schedule...")
def get_schedule(report_date_str):
    return fetch_todays_schedule(date.fromisoformat(report_date_str))


def build_report_bytes(games, triggers, report_date, odds):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    NAVY='#1B2A4A'; STEEL='#2E5F8A'; LIGHT_STEEL='#D0E4F5'
    FADE_RED='#C00000'; FADE_BG='#FFE7E7'; WATCH_AMB='#7D5A00'; WATCH_BG='#FFF3CC'
    GREEN_FG='#375623'; GREEN_BG='#E2EFDA'; AWAY_BG='#F0F5FB'; HOME_BG='#FFFFFF'; GRAY_TEXT='#666666'
    fb=wb.add_format({'bold':True,'font_size':16,'font_name':'Calibri','font_color':'white','bg_color':NAVY,'align':'center','valign':'vcenter'})
    fs=wb.add_format({'font_size':11,'font_name':'Calibri','italic':True,'font_color':LIGHT_STEEL,'bg_color':NAVY,'align':'center','valign':'vcenter'})
    fh=wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':'white','bg_color':STEEL,'align':'center','valign':'vcenter','border':1,'border_color':'#1A4A72'})
    fal=wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':STEEL,'bg_color':AWAY_BG,'align':'left','valign':'vcenter','left':2,'left_color':STEEL,'border':1,'border_color':'#B8C9DC','indent':1})
    fac=wb.add_format({'font_size':10,'font_name':'Calibri','bg_color':AWAY_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#B8C9DC'})
    fhlb=wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':'#333333','bg_color':HOME_BG,'align':'left','valign':'vcenter','left':2,'left_color':STEEL,'top':1,'top_color':'#B8C9DC','bottom':2,'bottom_color':STEEL,'right':1,'right_color':'#B8C9DC','indent':1})
    fhcb=wb.add_format({'font_size':10,'font_name':'Calibri','bg_color':HOME_BG,'align':'center','valign':'vcenter','top':1,'top_color':'#B8C9DC','bottom':2,'bottom_color':STEEL,'left':1,'left_color':'#B8C9DC','right':1,'right_color':'#B8C9DC'})
    ffa=wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':FADE_RED,'bg_color':FADE_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#E8AAAA'})
    ffhb=wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':FADE_RED,'bg_color':FADE_BG,'align':'center','valign':'vcenter','top':1,'top_color':'#E8AAAA','bottom':2,'bottom_color':STEEL,'left':1,'left_color':'#E8AAAA','right':1,'right_color':'#E8AAAA'})
    fba=wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':GREEN_FG,'bg_color':GREEN_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#A8D08D'})
    fbhb=wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':GREEN_FG,'bg_color':GREEN_BG,'align':'center','valign':'vcenter','top':1,'bottom':2,'bottom_color':STEEL,'left':1,'right':1,'border_color':'#A8D08D'})
    fwa=wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':WATCH_AMB,'bg_color':WATCH_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#D4B800'})
    fwhb=wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':WATCH_AMB,'bg_color':WATCH_BG,'align':'center','valign':'vcenter','top':1,'bottom':2,'bottom_color':STEEL,'left':1,'right':1,'border_color':'#D4B800'})
    fnpa=wb.add_format({'italic':True,'font_size':9,'font_name':'Calibri','font_color':GRAY_TEXT,'bg_color':AWAY_BG,'align':'center','valign':'vcenter','border':1,'border_color':'#B8C9DC'})
    fnph=wb.add_format({'italic':True,'font_size':9,'font_name':'Calibri','font_color':GRAY_TEXT,'bg_color':HOME_BG,'align':'center','valign':'vcenter','top':1,'bottom':2,'bottom_color':STEEL,'left':1,'right':1,'border_color':'#B8C9DC'})
    fng=wb.add_format({'italic':True,'font_size':11,'font_color':GRAY_TEXT,'align':'center','valign':'vcenter'})
    fgc=wb.add_format({'bold':True,'font_size':10,'font_name':'Calibri','font_color':NAVY,'bg_color':AWAY_BG,'align':'left','valign':'vcenter','border':1,'border_color':'#B8C9DC','indent':1})
    fgch=wb.add_format({'font_size':9,'font_name':'Calibri','font_color':GRAY_TEXT,'bg_color':HOME_BG,'align':'left','valign':'vcenter','top':1,'bottom':2,'bottom_color':STEEL,'left':1,'right':1,'border_color':'#B8C9DC'})
    fsl=wb.add_format({'bold':True,'font_size':11,'font_name':'Calibri','font_color':NAVY,'bg_color':'#EDF3FB','border':1,'border_color':'#B8C9DC','align':'left','valign':'vcenter','indent':1})
    fsv=wb.add_format({'bold':True,'font_size':14,'font_name':'Calibri','font_color':NAVY,'bg_color':'#EDF3FB','border':1,'border_color':'#B8C9DC','align':'center','valign':'vcenter'})
    ffl=wb.add_format({'bold':True,'font_size':11,'font_name':'Calibri','font_color':FADE_RED,'bg_color':FADE_BG,'border':1,'border_color':'#E8AAAA','align':'left','valign':'vcenter','indent':1})
    ffv=wb.add_format({'bold':True,'font_size':14,'font_name':'Calibri','font_color':FADE_RED,'bg_color':FADE_BG,'border':1,'border_color':'#E8AAAA','align':'center','valign':'vcenter'})
    fwl=wb.add_format({'bold':True,'font_size':11,'font_name':'Calibri','font_color':WATCH_AMB,'bg_color':WATCH_BG,'border':1,'border_color':'#D4B800','align':'left','valign':'vcenter','indent':1})
    fwv=wb.add_format({'bold':True,'font_size':14,'font_name':'Calibri','font_color':WATCH_AMB,'bg_color':WATCH_BG,'border':1,'border_color':'#D4B800','align':'center','valign':'vcenter'})
    fbl=wb.add_format({'bold':True,'font_size':11,'font_name':'Calibri','font_color':GREEN_FG,'bg_color':GREEN_BG,'border':1,'border_color':'#A8D08D','align':'left','valign':'vcenter','indent':1})
    fbv=wb.add_format({'bold':True,'font_size':14,'font_name':'Calibri','font_color':GREEN_FG,'bg_color':GREEN_BG,'border':1,'border_color':'#A8D08D','align':'center','valign':'vcenter'})

    NCOLS=5; CW=[32,8,26,10,50]; CH=['GAME','H/A','Team','Odds','Play  /  Scenario']

    def pfa(tl):
        if not tl: return fnpa
        v=tl[0]['verdict']
        return ffa if v=='CLEAR FADE' else (fba if v=='CLEAR BET' else fwa)
    def pfh(tl):
        if not tl: return fnph
        v=tl[0]['verdict']
        return ffhb if v=='CLEAR FADE' else (fbhb if v=='CLEAR BET' else fwhb)

    def write_tab(ws,label,vf,tc):
        ws.set_tab_color(tc)
        for ci,w in enumerate(CW): ws.set_column(ci,ci,w)
        ws.set_row(0,30); ws.set_row(1,18)
        ws.merge_range(0,0,0,NCOLS-1,f'MLB DAILY BETTING REPORT  -  {label}',fb)
        ws.merge_range(1,0,1,NCOLS-1,f'{report_date.strftime("%A, %B %d, %Y")}  |  Generated {datetime.now().strftime("%I:%M %p")}',fs)
        ws.set_row(2,20)
        for ci,h in enumerate(CH): ws.write(2,ci,h,fh)
        ws.freeze_panes(3,0)
        row=3; written=0
        for g in games:
            away,home=g['away_team'],g['home_team']
            atc,htc=title_case(away),title_case(home)
            at=[t for t in triggers if t['team']==away and t['opponent']==home and t['verdict']==vf]
            ht=[t for t in triggers if t['team']==home and t['opponent']==away and t['verdict']==vf]
            if not at and not ht: continue
            written+=1
            ws.set_row(row,20)
            ap=(at[0]['play']+'  |  '+'  |  '.join(f"#{t['scenario_id']} {t['scenario']}" for t in at)) if at else ''
            ws.write(row,0,f'{atc}  @  {htc}',fgc); ws.write(row,1,'AWAY',fac); ws.write(row,2,atc,fal)
            ws.write(row,3,fmt_line(odds.get(away)),fac); ws.write(row,4,ap,pfa(at)); row+=1
            ws.set_row(row,20)
            hp=(ht[0]['play']+'  |  '+'  |  '.join(f"#{t['scenario_id']} {t['scenario']}" for t in ht)) if ht else ''
            ws.write(row,0,'',fgch); ws.write(row,1,'HOME',fhcb); ws.write(row,2,htc,fhlb)
            ws.write(row,3,fmt_line(odds.get(home)),fhcb); ws.write(row,4,hp,pfh(ht)); row+=1
        if written==0: ws.merge_range(row,0,row,NCOLS-1,f'No {vf.title()} scenarios triggered today.',fng)
        return written

    w0=wb.add_worksheet('Green Clear Bet');   n0=write_tab(w0,'CLEAR BET','CLEAR BET','#00B050')
    w1=wb.add_worksheet('Red Clear Fade');    n1=write_tab(w1,'CLEAR FADE','CLEAR FADE','#FF0000')
    w2=wb.add_worksheet('Yellow Inconsistent');n2=write_tab(w2,'INCONSISTENT','INCONSISTENT','#FFD700')

    w3=wb.add_worksheet('Summary'); w3.set_tab_color('#1B2A4A')
    w3.set_column(0,0,35); w3.set_column(1,1,18); w3.set_column(2,2,35); w3.set_column(3,3,18)
    w3.set_row(0,30); w3.set_row(1,18)
    w3.merge_range(0,0,0,3,'MLB DAILY BETTING REPORT  -  SUMMARY',fb)
    w3.merge_range(1,0,1,3,f'{report_date.strftime("%A, %B %d, %Y")}  |  Generated {datetime.now().strftime("%I:%M %p")}',fs)
    nbt=sum(1 for t in triggers if t['verdict']=='CLEAR BET')
    nft=sum(1 for t in triggers if t['verdict']=='CLEAR FADE')
    nit=sum(1 for t in triggers if t['verdict']=='INCONSISTENT')
    w3.set_row(3,28); w3.set_row(4,28); w3.set_row(5,28)
    w3.write(3,0,'Total Games Today',fsl); w3.write(3,1,len(games),fsv)
    w3.write(3,2,'Games With Triggers',fsl); w3.write(3,3,n0+n1+n2,fsv)
    w3.write(4,0,'Clear Bet Triggers',fbl); w3.write(4,1,nbt,fbv)
    w3.write(4,2,'Clear Fade Triggers',ffl); w3.write(4,3,nft,ffv)
    w3.write(5,0,'Inconsistent Triggers',fwl); w3.write(5,1,nit,fwv)
    w3.merge_range(6,0,6,3,'TOP PLAYS TODAY',fh); w3.set_row(7,20)
    for ci,h in enumerate(['Team','Odds','Play','Scenario']): w3.write(7,ci,h,fh)
    sr=8
    for t in [x for x in triggers if x['verdict']=='CLEAR BET']:
        w3.write(sr,0,title_case(t['team']),fbl); w3.write(sr,1,fmt_line(t['line']),fbv)
        w3.write(sr,2,t['play'],fbl); w3.write(sr,3,f"#{t['scenario_id']} {t['scenario']}",fbl); sr+=1
    for t in [x for x in triggers if x['verdict']=='CLEAR FADE']:
        w3.write(sr,0,title_case(t['team']),ffl); w3.write(sr,1,fmt_line(t['line']),ffv)
        w3.write(sr,2,t['play'],ffl); w3.write(sr,3,f"#{t['scenario_id']} {t['scenario']}",ffl); sr+=1
    for t in [x for x in triggers if x['verdict']=='INCONSISTENT']:
        w3.write(sr,0,title_case(t['team']),fwl); w3.write(sr,1,fmt_line(t['line']),fwv)
        w3.write(sr,2,t['play'],fwl); w3.write(sr,3,f"#{t['scenario_id']} {t['scenario']}",fwl); sr+=1
    w3.freeze_panes(8,0)

    w4=wb.add_worksheet('Results Tracker'); w4.set_tab_color('#375623')
    w4.set_column(0,0,12); w4.set_column(1,1,26); w4.set_column(2,2,10); w4.set_column(3,3,10)
    w4.set_column(4,4,14); w4.set_column(5,5,40); w4.set_column(6,6,14); w4.set_column(7,7,14); w4.set_column(8,8,16)
    fth=wb.add_format({'bold':True,'font_color':'white','bg_color':'#375623','align':'center','valign':'vcenter','border':1,'font_name':'Calibri'})
    ftc=wb.add_format({'align':'center','valign':'vcenter','border':1,'border_color':'#C6EFCE','font_name':'Calibri'})
    ftl=wb.add_format({'align':'left','valign':'vcenter','border':1,'border_color':'#C6EFCE','font_name':'Calibri','indent':1})
    fti=wb.add_format({'bold':True,'align':'center','valign':'vcenter','border':2,'border_color':'#375623','bg_color':'#EAF4E8','font_name':'Calibri'})
    ftt=wb.add_format({'bold':True,'bg_color':'#375623','font_color':'white','align':'center','border':1,'font_name':'Calibri'})
    w4.set_row(0,30); w4.set_row(1,18)
    w4.merge_range(0,0,0,8,'RESULTS TRACKER  -  Enter W or L after each game',fb)
    w4.merge_range(1,0,1,8,'Enter W or L in the Result column. Net P/L calculates automatically.',fs)
    w4.set_row(2,30)
    for ci,h in enumerate(['Date','Team','H/A','Odds','Play','Scenario','Type','Result (W/L)','Net P/L ($100)']): w4.write(2,ci,h,fth)
    tr=3
    for t in triggers:
        line=t['line']; er=tr+1
        w4.write(tr,0,report_date.strftime('%Y-%m-%d'),ftc); w4.write(tr,1,title_case(t['team']),ftl)
        w4.write(tr,2,t['home_away'].upper(),ftc); w4.write(tr,3,fmt_line(line),ftc)
        w4.write(tr,4,t['play'],ftl); w4.write(tr,5,f"#{t['scenario_id']} {t['scenario']}",ftl)
        w4.write(tr,6,t['verdict'],ftc); w4.write(tr,7,'',fti)
        if line is not None and isinstance(line,int):
            pf=f'=IF(H{er}="W",{line},IF(H{er}="L",-100,""))' if line>0 else f'=IF(H{er}="W",ROUND(100/ABS({line})*100,2),IF(H{er}="L",-100,""))'
            w4.write_formula(tr,8,pf,ftc)
        else: w4.write(tr,8,'',ftc)
        tr+=1
    if tr>3:
        w4.merge_range(tr,0,tr,7,'TOTAL NET P/L',ftt)
        w4.write_formula(tr,8,f'=SUMIF(H4:H{tr},"W",I4:I{tr})+SUMIF(H4:H{tr},"L",I4:I{tr})',ftt)
    w4.freeze_panes(3,0); w4.autofilter(2,0,tr,8)

    # Scenario Performance tab
    w5=wb.add_worksheet('Scenario Performance'); w5.set_tab_color('#1F3864')
    w5.set_column(0,0,6); w5.set_column(1,1,40); w5.set_column(2,2,14)
    w5.set_column(3,3,8); w5.set_column(4,4,8); w5.set_column(5,5,10)
    w5.set_column(6,6,10); w5.set_column(7,7,14)
    fpb=wb.add_format({"bold":True,"font_size":14,"font_name":"Calibri","font_color":"white","bg_color":"#1F3864","align":"center","valign":"vcenter"})
    fps=wb.add_format({"font_size":10,"font_name":"Calibri","italic":True,"font_color":"#D0E4F5","bg_color":"#1F3864","align":"center","valign":"vcenter"})
    fph=wb.add_format({"bold":True,"font_size":10,"font_name":"Calibri","font_color":"white","bg_color":"#2E5F8A","align":"center","valign":"vcenter","border":1})
    fps2=wb.add_format({"bold":True,"font_size":10,"font_name":"Calibri","font_color":"#1F3864","bg_color":"#EDF3FB","align":"center","valign":"vcenter","border":1})
    fpn=wb.add_format({"font_size":10,"font_name":"Calibri","font_color":"#1F3864","bg_color":"#EDF3FB","align":"left","valign":"vcenter","border":1,"indent":1})
    fpbet=wb.add_format({"font_size":9,"bold":True,"font_color":"#375623","bg_color":"#E2EFDA","align":"center","valign":"vcenter","border":1,"font_name":"Calibri"})
    fpfade=wb.add_format({"font_size":9,"bold":True,"font_color":"#C00000","bg_color":"#FFE7E7","align":"center","valign":"vcenter","border":1,"font_name":"Calibri"})
    fpinc=wb.add_format({"font_size":9,"bold":True,"font_color":"#7D5A00","bg_color":"#FFF3CC","align":"center","valign":"vcenter","border":1,"font_name":"Calibri"})
    fpnum=wb.add_format({"font_size":10,"bg_color":"#EDF3FB","align":"center","valign":"vcenter","border":1,"font_name":"Calibri"})
    fppct=wb.add_format({"font_size":10,"num_format":"0.000","bg_color":"#EDF3FB","align":"center","valign":"vcenter","border":1,"font_name":"Calibri"})
    fpmny=wb.add_format({"font_size":10,"num_format":"0,##0.00","bg_color":"#EDF3FB","align":"center","valign":"vcenter","border":1,"font_name":"Calibri"})
    fptot=wb.add_format({"bold":True,"font_size":10,"font_color":"white","bg_color":"#1F3864","align":"center","valign":"vcenter","border":1,"font_name":"Calibri"})
    w5.set_row(0,28); w5.set_row(1,16)
    w5.merge_range(0,0,0,7,"SCENARIO PERFORMANCE TRACKER  -  Season Cumulative",fpb)
    w5.merge_range(1,0,1,7,"Updates automatically as you enter results in the Results Tracker tab.",fps)
    w5.set_row(2,20)
    for ci,h in enumerate(["#","Scenario Name","Classification","W","L","Total","Win%","Net P/L"]):
        w5.write(2,ci,h,fph)
    w5.freeze_panes(3,0)
    tsheet="'Results Tracker'"; tend=max(tr+500,1000)
    pr=3
    from daily_report import SCENARIO_DEFS
    for sid,sname,verdict,_ in SCENARIO_DEFS:
        sid_str=f"#{sid} {sname}"
        vfmt=fpbet if verdict=="CLEAR BET" else (fpfade if verdict=="CLEAR FADE" else fpinc)
        vlabel="CLEAR BET" if verdict=="CLEAR BET" else ("CLEAR FADE" if verdict=="CLEAR FADE" else "INCONSISTENT")
        er=pr+1
        w5.write(pr,0,sid,fps2); w5.write(pr,1,sname,fpn); w5.write(pr,2,vlabel,vfmt)
        w5.write_formula(pr,3,'=COUNTIFS('+sc+',"'+sid_str+'",'+rc+',"W")',fpnum)
        w5.write_formula(pr,4,'=COUNTIFS('+sc+',"'+sid_str+'",'+rc+',"L")',fpnum)
        w5.write_formula(pr,5,f'=D{er}+E{er}',fpnum)
        w5.write_formula(pr,6,f'=IF(F{er}>0,D{er}/F{er},"")',fppct)
        w5.write_formula(pr,7,'=SUMPRODUCT(('+sc+'="'+sid_str+'")*ISNUMBER(MATCH('+rc+',{"W","L"},0))*('+pc+'))',fpmny)
        pr+=1
    w5.set_row(pr,20)
    for ci in range(8): w5.write(pr,ci,"" if ci not in [1] else "SEASON TOTALS",fptot)
    w5.write_formula(pr,3,f"=SUM(D4:D{pr})",fptot); w5.write_formula(pr,4,f"=SUM(E4:E{pr})",fptot)
    w5.write_formula(pr,5,f"=SUM(F4:F{pr})",fptot)
    w5.write_formula(pr,7,f"=SUM(H4:H{pr})",fptot)

    wb.close(); output.seek(0)
    return output.getvalue(), n0+n1+n2, n1, n2


# ── MAIN UI ───────────────────────────────────────────────────────
st.title("⚾ MLB Daily Betting Report")
st.markdown("---")

col1, col2 = st.columns([2,4])
with col1:
    report_date = st.date_input("Report Date", value=date.today())
report_date_str = report_date.isoformat()

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

st.subheader("📋 Enter Today's Moneylines")
st.caption("Enter the moneyline for each team (e.g. 130 for +130, -150 for -150). Leave blank if unavailable.")

odds_data = [{'Away Team': g['away_team'], 'Home Team': g['home_team'], 'Away Line': None, 'Home Line': None} for g in games]
odds_df = pd.DataFrame(odds_data)
edited = st.data_editor(
    odds_df,
    column_config={
        'Away Team': st.column_config.TextColumn('Away Team', disabled=True, width='large'),
        'Home Team': st.column_config.TextColumn('Home Team', disabled=True, width='large'),
        'Away Line': st.column_config.NumberColumn('Away Line', help='e.g. 130 or -150', width='medium'),
        'Home Line': st.column_config.NumberColumn('Home Line', help='e.g. -130 or 150', width='medium'),
    },
    hide_index=True, use_container_width=True,
)

st.markdown("---")

if st.button("⚾ Generate Daily Report", type="primary", use_container_width=True):
    odds = {}
    for _, row in edited.iterrows():
        if row['Away Line'] is not None and str(row['Away Line']) not in ['','nan']:
            try: odds[str(row['Away Team']).upper()] = int(row['Away Line'])
            except: pass
        if row['Home Line'] is not None and str(row['Home Line']) not in ['','nan']:
            try: odds[str(row['Home Team']).upper()] = int(row['Home Line'])
            except: pass

    if not odds:
        st.warning("Please enter at least some moneylines before generating the report.")
        st.stop()

    with st.spinner("Running scenario analysis..."):
        opp_streaks = {}; opp_road_wpct = {}
        for team in API_TO_CANONICAL.values():
            tdf = enriched[(enriched['team']==team) & (enriched['date'].dt.date < report_date)]
            if not tdf.empty:
                last = tdf.iloc[-1]
                sb = last['streak_before']; res = last['result']
                opp_streaks[team] = (sb+1 if sb>=0 else 1) if res=='W' else (sb-1 if sb<=0 else -1)
                road = tdf[tdf['home_away']=='away']
                if not road.empty:
                    rw=(road['result']=='W').sum(); rl=(road['result']=='L').sum()
                    opp_road_wpct[team] = rw/(rw+rl) if (rw+rl)>0 else None

        all_triggers = []
        for g in games:
            away, home = g['away_team'], g['home_team']
            away_state = get_team_state(enriched, away, report_date)
            home_state = get_team_state(enriched, home, report_date)
            away_row = build_game_row(away_state, 'away', home, odds.get(away))
            home_row = build_game_row(home_state, 'home', away, odds.get(home))
            all_triggers.extend(check_scenarios([away_row, home_row], opp_streaks, opp_road_wpct))

        excel_bytes, n_total, n_fade, n_inc = build_report_bytes(games, all_triggers, report_date, odds)

    st.markdown("---")
    st.subheader("📊 Results")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Games", len(games))
    m2.metric("🟢 Clear Bet", sum(1 for t in all_triggers if t['verdict']=='CLEAR BET'))
    m3.metric("🔴 Clear Fade", sum(1 for t in all_triggers if t['verdict']=='CLEAR FADE'))
    m4.metric("🟡 Inconsistent", sum(1 for t in all_triggers if t['verdict']=='INCONSISTENT'))

    for section_label, section_verdict in [
        ("🟢 Clear Bet Plays", "CLEAR BET"),
        ("🔴 Clear Fade Plays", "CLEAR FADE"),
        ("🟡 Inconsistent Plays", "INCONSISTENT"),
    ]:
        section_triggers = [t for t in all_triggers if t['verdict']==section_verdict]
        if not section_triggers: continue
        st.markdown(f"### {section_label}")
        game_rows = []
        seen = set()
        for g in games:
            away, home = g['away_team'], g['home_team']
            key = f"{away}@{home}"
            if key in seen: continue
            at = [t for t in section_triggers if t['team']==away and t['opponent']==home]
            ht = [t for t in section_triggers if t['team']==home and t['opponent']==away]
            if not at and not ht: continue
            seen.add(key)
            matchup = f"{title_case(away)} @ {title_case(home)}"
            game_rows.append({'GAME': matchup, 'H/A': 'AWAY', 'Team': title_case(away),
                'Odds': fmt_line(odds.get(away)),
                'Play': at[0]['play'] if at else '',
                'Scenario': f"#{at[0]['scenario_id']} {at[0]['scenario']}" if at else ''})
            game_rows.append({'GAME': '', 'H/A': 'HOME', 'Team': title_case(home),
                'Odds': fmt_line(odds.get(home)),
                'Play': ht[0]['play'] if ht else '',
                'Scenario': f"#{ht[0]['scenario_id']} {ht[0]['scenario']}" if ht else ''})
        if game_rows:
            st.dataframe(pd.DataFrame(game_rows), use_container_width=True, hide_index=True,
                column_config={
                    'GAME': st.column_config.TextColumn('GAME', width='large'),
                    'H/A': st.column_config.TextColumn('H/A', width='small'),
                    'Team': st.column_config.TextColumn('Team', width='medium'),
                    'Odds': st.column_config.TextColumn('Odds', width='small'),
                    'Play': st.column_config.TextColumn('Play', width='medium'),
                    'Scenario': st.column_config.TextColumn('Scenario', width='large'),
                })

    st.markdown("---")
    fname = f'MLB_Daily_Report_{report_date.strftime("%Y-%m-%d")}.xlsx'
    st.download_button(label="📥 Download Excel Report", data=excel_bytes, file_name=fname,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True, type="primary")
