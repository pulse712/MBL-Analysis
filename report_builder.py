"""Shared Excel report builder for Streamlit downloads (in-memory bytes)."""
import io
from datetime import datetime

import xlsxwriter

from daily_report import (
    SCENARIO_DEFS,
    fmt_line,
    numeric_line,
    pl_line_for_trigger,
    title_case,
)


def build_report_bytes(games, triggers, report_date, odds, cumulative_df=None):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True, 'calc_on_load': True})
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
    w4.set_column(9, 9, None, None, {'hidden': True})
    w4.set_column(10, 10, None, None, {'hidden': True})
    fth=wb.add_format({'bold':True,'font_color':'white','bg_color':'#375623','align':'center','valign':'vcenter','border':1,'font_name':'Calibri'})
    ftc=wb.add_format({'align':'center','valign':'vcenter','border':1,'border_color':'#C6EFCE','font_name':'Calibri'})
    ftl=wb.add_format({'align':'left','valign':'vcenter','border':1,'border_color':'#C6EFCE','font_name':'Calibri','indent':1})
    fti=wb.add_format({'bold':True,'align':'center','valign':'vcenter','border':2,'border_color':'#375623','bg_color':'#EAF4E8','font_name':'Calibri'})
    ftt=wb.add_format({'bold':True,'bg_color':'#375623','font_color':'white','align':'center','border':1,'font_name':'Calibri'})
    w4.set_row(0,30); w4.set_row(1,18)
    w4.merge_range(0,0,0,8,'RESULTS TRACKER  -  Enter W or L after each game',fb)
    w4.merge_range(1,0,1,8,'Enter W or L in the Result column. Net P/L calculates automatically.',fs)
    w4.set_row(2,30)
    for ci,h in enumerate(['Date','Team','H/A','Odds','Play','Scenario','Type','Result (W/L)','Net P/L ($100)','PayoutLine','Opponent']):
        w4.write(2,ci,h,fth)
    tr=3
    for t in triggers:
        line=t['line']
        pl_line = numeric_line(pl_line_for_trigger(t))
        er=tr+1
        w4.write(tr,0,report_date.strftime('%Y-%m-%d'),ftc); w4.write(tr,1,title_case(t['team']),ftl)
        w4.write(tr,2,t['home_away'].upper(),ftc); w4.write(tr,3,fmt_line(line),ftc)
        w4.write(tr,4,t['play'],ftl); w4.write(tr,5,f"#{t['scenario_id']} {t['scenario']}",ftl)
        w4.write(tr,6,t['verdict'],ftc); w4.write(tr,7,'',fti)
        if pl_line is not None:
            pf=f'=IF(H{er}="W",{pl_line},IF(H{er}="L",-100,""))' if pl_line>0 else f'=IF(H{er}="W",ROUND(100/ABS({pl_line})*100,2),IF(H{er}="L",-100,""))'
            w4.write_formula(tr,8,pf,ftc)
        else: w4.write(tr,8,'',ftc)
        if pl_line is not None:
            w4.write(tr, 9, pl_line, ftc)
        w4.write(tr, 10, str(t.get('opponent', '') or '').strip().upper(), ftc)
        tr+=1
    if tr>3:
        w4.merge_range(tr,0,tr,7,'TOTAL NET P/L',ftt)
        w4.write_formula(tr,8,f'=SUMIF(H4:H{tr},"W",I4:I{tr})+SUMIF(H4:H{tr},"L",I4:I{tr})',ftt)
    w4.freeze_panes(3,0); w4.autofilter(2,0,tr,8)

    # Optional cumulative results sheet (prior days only — today lives on Results Tracker)
    cum_data_end = 3
    report_day_str = report_date.strftime('%Y-%m-%d')
    hist_rows = []
    if cumulative_df is not None and not cumulative_df.empty:
        for _, row in cumulative_df.iterrows():
            if str(row.get('Date', ''))[:10] < report_day_str:
                hist_rows.append(row)
    has_cumulative = len(hist_rows) > 0
    if has_cumulative:
        w_cum = wb.add_worksheet('Cumulative Results')
        w_cum.hide()
        w_cum.set_column(0, 0, 12); w_cum.set_column(1, 1, 26); w_cum.set_column(5, 5, 40)
        w_cum.set_column(7, 7, 14); w_cum.set_column(8, 8, 16)
        w_cum.set_column(9, 9, None, None, {'hidden': True})
        w_cum.set_column(10, 10, None, None, {'hidden': True})
        for ci, h in enumerate(['Date', 'Team', 'H/A', 'Odds', 'Play', 'Scenario', 'Type', 'Result (W/L)', 'Net P/L ($100)', 'PayoutLine', 'Opponent']):
            w_cum.write(2, ci, h, fth)
        cum_row = 3
        for row in hist_rows:
            er = cum_row + 1
            w_cum.write(cum_row, 0, str(row.get('Date', '')), ftc)
            w_cum.write(cum_row, 1, str(row.get('Team', '')), ftl)
            w_cum.write(cum_row, 2, str(row.get('H/A', '')), ftc)
            w_cum.write(cum_row, 3, str(row.get('Odds', '')), ftc)
            w_cum.write(cum_row, 4, str(row.get('Play', '')), ftl)
            w_cum.write(cum_row, 5, str(row.get('Scenario', '')), ftl)
            w_cum.write(cum_row, 6, str(row.get('Type', '')), ftc)
            res = str(row.get('Result', '') or '').strip().upper()
            w_cum.write(cum_row, 7, res if res in ('W', 'L') else '', fti if res in ('W', 'L') else ftc)
            payout = numeric_line(row.get('PayoutLine'))
            if payout is None:
                payout = numeric_line(str(row.get('Odds', '')).replace('+', ''))
            if payout is not None:
                pf = (f'=IF(H{er}="W",{payout},IF(H{er}="L",-100,""))' if payout > 0
                      else f'=IF(H{er}="W",ROUND(100/ABS({payout})*100,2),IF(H{er}="L",-100,""))')
                w_cum.write_formula(cum_row, 8, pf, ftc)
                w_cum.write(cum_row, 9, payout, ftc)
            else:
                w_cum.write(cum_row, 8, '', ftc)
            w_cum.write(cum_row, 10, str(row.get('Opponent', row.get('_opponent', '')) or '').strip().upper(), ftc)
            cum_row += 1
        cum_data_end = cum_row

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
    fppct=wb.add_format({"font_size":10,"num_format":"0.0%","bg_color":"#EDF3FB","align":"center","valign":"vcenter","border":1,"font_name":"Calibri"})
    fpmny=wb.add_format({"font_size":10,"num_format":"$#,##0.00","bg_color":"#EDF3FB","align":"center","valign":"vcenter","border":1,"font_name":"Calibri"})
    fptot=wb.add_format({"bold":True,"font_size":10,"font_color":"white","bg_color":"#1F3864","align":"center","valign":"vcenter","border":1,"font_name":"Calibri"})
    fpnote=wb.add_format({"font_size":9,"font_name":"Calibri","italic":True,"font_color":"#7D5A00","bg_color":"#FFF9E6","align":"center","valign":"vcenter","border":1,"text_wrap":True})

    w5.set_row(0,28); w5.set_row(1,16); w5.set_row(2,22)
    tr_end = max(tr + 50, 200)
    tr_sc = "'Results Tracker'!F$4:F$" + str(tr_end)
    tr_rc = "'Results Tracker'!H$4:H$" + str(tr_end)
    tr_pc = "'Results Tracker'!I$4:I$" + str(tr_end)
    if has_cumulative:
        w5.merge_range(0,0,0,7,"SCENARIO PERFORMANCE TRACKER  -  Season Cumulative",fpb)
        w5.merge_range(1,0,1,7,"Prior days from uploaded master + today from Results Tracker (enter W/L there).",fps)
        w5.merge_range(2,0,2,7,"Download Master_Results.xlsx below to persist season totals across days.",fpnote)
        cum_end = max(cum_data_end + 50, 200)
        cum_sc = "'Cumulative Results'!F$4:F$" + str(cum_end)
        cum_rc = "'Cumulative Results'!H$4:H$" + str(cum_end)
        cum_pc = "'Cumulative Results'!I$4:I$" + str(cum_end)
    else:
        w5.merge_range(0,0,0,7,"SCENARIO PERFORMANCE TRACKER  -  Today's Results Only",fpb)
        w5.merge_range(1,0,1,7,"Upload Master_Results.xlsx in the app before generating to see season-long stats here.",fps)
        w5.merge_range(2,0,2,7,"To build cumulative stats: upload master file → generate → download updated Master_Results.xlsx.",fpnote)
    w5.set_row(3,20)
    for ci,h in enumerate(["#","Scenario Name","Classification","W","L","Total","Win%","Net P/L"]):
        w5.write(3,ci,h,fph)
    w5.freeze_panes(4,0)

    pr=4
    # SCENARIO_DEFS imported at module level
    for sid,sname,verdict,_ in SCENARIO_DEFS:
        sid_str=f"#{sid} {sname}"
        vfmt=fpbet if verdict=="CLEAR BET" else (fpfade if verdict=="CLEAR FADE" else fpinc)
        vlabel="CLEAR BET" if verdict=="CLEAR BET" else ("CLEAR FADE" if verdict=="CLEAR FADE" else "INCONSISTENT")
        er=pr+1
        w5.write(pr,0,sid,fps2); w5.write(pr,1,sname,fpn); w5.write(pr,2,vlabel,vfmt)
        if has_cumulative:
            w5.write_formula(pr,3,f'=IFERROR(COUNTIFS({cum_sc},"{sid_str}",{cum_rc},"W"),0)+IFERROR(COUNTIFS({tr_sc},"{sid_str}",{tr_rc},"W"),0)',fpnum)
            w5.write_formula(pr,4,f'=IFERROR(COUNTIFS({cum_sc},"{sid_str}",{cum_rc},"L"),0)+IFERROR(COUNTIFS({tr_sc},"{sid_str}",{tr_rc},"L"),0)',fpnum)
            w5.write_formula(pr,7,f'=IFERROR(SUMIFS({cum_pc},{cum_sc},"{sid_str}"),0)+IFERROR(SUMIFS({tr_pc},{tr_sc},"{sid_str}"),0)',fpmny)
        else:
            w5.write_formula(pr,3,f'=IFERROR(COUNTIFS({tr_sc},"{sid_str}",{tr_rc},"W"),0)',fpnum)
            w5.write_formula(pr,4,f'=IFERROR(COUNTIFS({tr_sc},"{sid_str}",{tr_rc},"L"),0)',fpnum)
            w5.write_formula(pr,7,f'=IFERROR(SUMIFS({tr_pc},{tr_sc},"{sid_str}"),0)',fpmny)
        w5.write_formula(pr,5,f'=D{er}+E{er}',fpnum)
        w5.write_formula(pr,6,f'=IF(F{er}>0,D{er}/F{er},"")',fppct)
        pr+=1
    w5.set_row(pr,20)
    for ci in range(8): w5.write(pr,ci,"" if ci not in [1] else ("SEASON TOTALS" if has_cumulative else "TODAY TOTALS"),fptot)
    w5.write_formula(pr,3,f"=SUM(D5:D{pr})",fptot); w5.write_formula(pr,4,f"=SUM(E5:E{pr})",fptot)
    w5.write_formula(pr,5,f"=SUM(F5:F{pr})",fptot)
    w5.write_formula(pr,6,f'=IF(F{pr+1}>0,D{pr+1}/F{pr+1},"")',fptot)
    w5.write_formula(pr,7,f"=SUM(H5:H{pr})",fptot)

    wb.close(); output.seek(0)
    return output.getvalue(), n0+n1+n2, n1, n2
