"""
Master Results Manager
======================
Manages the cumulative Master_Results.xlsx file that stores all betting results
across the entire season for scenario performance tracking.
"""

import io
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
from datetime import datetime


MASTER_FILE = 'Master_Results.xlsx'
RESULTS_COLUMNS = ['Date', 'Team', 'H/A', 'Odds', 'Play', 'Scenario', 'Type', 'Result', 'Net P/L']


def _detect_header_row(probe_df):
    """Return the row index that contains the 'Date' column header."""
    for i in range(min(6, len(probe_df))):
        row_vals = [str(v).strip() for v in probe_df.iloc[i].tolist()]
        if 'Date' in row_vals:
            return i
    return 2


def _find_results_sheet(sheet_names):
    """
    Pick the sheet that holds trackable results.
    Accepts Master_Results.xlsx and daily reports (Results Tracker tab).
    """
    if 'Master Results' in sheet_names:
        return 'Master Results', 'Master_Results.xlsx'
    for name in sheet_names:
        if name == 'Results Tracker' or 'Results Tracker' in name:
            return name, 'daily report'
    return None, None


def _normalize_results_df(df):
    """Normalize uploaded results to the standard 9-column layout."""
    # Read up to 10 columns — col 10 is the hidden PayoutLine (payout line for P/L)
    has_payout_col = len(df.columns) >= 10
    if has_payout_col:
        df = df.iloc[:, :10].copy()
        df.columns = RESULTS_COLUMNS + ['PayoutLine']
    else:
        df = df.iloc[:, :9].copy()
        df.columns = RESULTS_COLUMNS
    df = df[df['Date'].notna()]
    # Normalize Date to YYYY-MM-DD string regardless of how pandas read it
    def _fmt_date(v):
        try:
            return pd.to_datetime(v).strftime('%Y-%m-%d')
        except Exception:
            return str(v).strip()
    df['Date'] = df['Date'].apply(_fmt_date)
    date_str = df['Date'].astype(str).str.strip().str.upper()
    team_str = df['Team'].astype(str).str.strip().str.upper()
    df = df[~date_str.str.contains('TOTAL', na=False)]
    df = df[~team_str.str.contains('TOTAL', na=False)]
    df = df[date_str.str.match(r'\d{4}-\d{2}-\d{2}', na=False)]
    df['Result'] = df['Result'].astype(str).str.strip().str.upper()
    df.loc[~df['Result'].isin(['W', 'L']), 'Result'] = ''
    return df.reset_index(drop=True)


def parse_results_upload(file_bytes):
    """
    Load trackable results from an uploaded Excel file.

    Accepts:
    - Master_Results.xlsx (Master Results sheet)
    - MLB_Daily_Report_*.xlsx (Results Tracker sheet, with W/L entered)

    Returns:
        (DataFrame, source_label, sheet_name)
    """
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet, source = _find_results_sheet(xl.sheet_names)

    if sheet is None:
        for name in xl.sheet_names:
            probe = pd.read_excel(io.BytesIO(file_bytes), sheet_name=name, header=None, nrows=6)
            if 'Date' in [str(v).strip() for v in probe.iloc[_detect_header_row(probe)].tolist()]:
                sheet, source = name, f'sheet "{name}"'
                break

    if sheet is None:
        raise ValueError(
            "No results data found. Upload Master_Results.xlsx (recommended) or a saved "
            "daily report whose Results Tracker tab has W/L results entered."
        )

    probe = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None, nrows=6)
    skip = _detect_header_row(probe)
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, skiprows=skip)
    df = _normalize_results_df(df)

    if df.empty:
        raise ValueError(
            f"Found sheet '{sheet}' but no result rows. "
            "Open the daily report, enter W/L on the Results Tracker tab, save, then upload."
        )

    return df, source, sheet


def initialize_master_results():
    """
    Create a new Master_Results.xlsx file if it doesn't exist.
    This file has the same structure as the Results Tracker tab.
    """
    if os.path.exists(MASTER_FILE):
        print(f'Master results file already exists: {MASTER_FILE}')
        return

    # Create a new workbook with the Results Tracker structure
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = 'Master Results'

    # Define styles
    header_fill = PatternFill(start_color='375623', end_color='375623', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF', name='Calibri')
    header_alignment = Alignment(horizontal='center', vertical='center')
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Set column widths
    ws.column_dimensions['A'].width = 12   # Date
    ws.column_dimensions['B'].width = 26   # Team
    ws.column_dimensions['C'].width = 10   # H/A
    ws.column_dimensions['D'].width = 10   # Odds
    ws.column_dimensions['E'].width = 14   # Play
    ws.column_dimensions['F'].width = 40   # Scenario
    ws.column_dimensions['G'].width = 14   # Type
    ws.column_dimensions['H'].width = 14   # Result
    ws.column_dimensions['I'].width = 16   # Net P/L

    # Title row
    ws.merge_cells('A1:I1')
    title_cell = ws['A1']
    title_cell.value = 'MASTER RESULTS TRACKER — All Season Results'
    title_cell.font = Font(bold=True, size=16, color='FFFFFF', name='Calibri')
    title_cell.fill = PatternFill(start_color='1B2A4A', end_color='1B2A4A', fill_type='solid')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 30

    # Subtitle row
    ws.merge_cells('A2:I2')
    subtitle_cell = ws['A2']
    subtitle_cell.value = 'Copy results from daily reports into this file to build cumulative season statistics'
    subtitle_cell.font = Font(italic=True, size=10, color='D0E4F5', name='Calibri')
    subtitle_cell.fill = PatternFill(start_color='1B2A4A', end_color='1B2A4A', fill_type='solid')
    subtitle_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18

    # Header row
    headers = ['Date', 'Team', 'H/A', 'Odds', 'Play', 'Scenario', 'Type', 'Result (W/L)', 'Net P/L ($100)']
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    ws.row_dimensions[3].height = 30

    # Add data validation for Result column (H)
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1='"W,L"', allow_blank=True)
    dv.error = 'Please enter W or L'
    dv.errorTitle = 'Invalid Entry'
    ws.add_data_validation(dv)
    dv.add('H4:H10000')

    # Freeze panes
    ws.freeze_panes = 'A4'

    wb.save(MASTER_FILE)
    print(f'✓ Created new master results file: {MASTER_FILE}')


def append_to_master_results(triggers, report_date):
    """
    Append today's triggers to the Master_Results.xlsx file.
    This allows the client to track results across multiple days.
    
    Args:
        triggers: List of trigger dictionaries from today's report
        report_date: Date object for today's report
    """
    if not os.path.exists(MASTER_FILE):
        initialize_master_results()

    # Load existing data
    wb = load_workbook(MASTER_FILE)
    ws = wb.active

    # Find the next empty row
    next_row = 4  # Start after header (row 3)
    while ws.cell(row=next_row, column=1).value is not None:
        next_row += 1

    # Define cell styles
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    cell_alignment = Alignment(horizontal='center', vertical='center')
    left_alignment = Alignment(horizontal='left', vertical='center', indent=1)
    thin_border = Border(
        left=Side(style='thin', color='C6EFCE'),
        right=Side(style='thin', color='C6EFCE'),
        top=Side(style='thin', color='C6EFCE'),
        bottom=Side(style='thin', color='C6EFCE')
    )
    input_fill = PatternFill(start_color='EAF4E8', end_color='EAF4E8', fill_type='solid')
    input_font = Font(bold=True, name='Calibri')
    input_border = Border(
        left=Side(style='medium', color='375623'),
        right=Side(style='medium', color='375623'),
        top=Side(style='medium', color='375623'),
        bottom=Side(style='medium', color='375623')
    )

    # Add each trigger as a row
    for trigger in triggers:
        row = next_row
        
        # Date
        cell = ws.cell(row=row, column=1)
        cell.value = report_date.strftime('%Y-%m-%d')
        cell.alignment = cell_alignment
        cell.border = thin_border

        # Team
        cell = ws.cell(row=row, column=2)
        cell.value = trigger['team'].title()
        cell.alignment = left_alignment
        cell.border = thin_border

        # H/A
        cell = ws.cell(row=row, column=3)
        cell.value = trigger['home_away'].upper()
        cell.alignment = cell_alignment
        cell.border = thin_border

        # Odds
        cell = ws.cell(row=row, column=4)
        line = trigger['line']
        if line is not None:
            cell.value = f'+{line}' if line > 0 else str(line)
        else:
            cell.value = 'N/A'
        cell.alignment = cell_alignment
        cell.border = thin_border

        # Play
        cell = ws.cell(row=row, column=5)
        cell.value = trigger['play']
        cell.alignment = left_alignment
        cell.border = thin_border

        # Scenario
        cell = ws.cell(row=row, column=6)
        cell.value = f"#{trigger['scenario_id']} {trigger['scenario']}"
        cell.alignment = left_alignment
        cell.border = thin_border

        # Type
        cell = ws.cell(row=row, column=7)
        cell.value = trigger['verdict']
        cell.alignment = cell_alignment
        cell.border = thin_border

        # Result (W/L) - empty for client to fill
        cell = ws.cell(row=row, column=8)
        cell.value = ''
        cell.alignment = cell_alignment
        cell.fill = input_fill
        cell.font = input_font
        cell.border = input_border

        # Net P/L - formula
        cell = ws.cell(row=row, column=9)
        line = trigger['line']
        if line is not None and isinstance(line, int):
            if line > 0:
                # Dog: if win get +line, if lose get -100
                formula = f'=IF(H{row}="W",{line},IF(H{row}="L",-100,""))'
            else:
                # Favorite: if win get 100/|line|*100, if lose get -100
                formula = f'=IF(H{row}="W",ROUND(100/ABS({line})*100,2),IF(H{row}="L",-100,""))'
            cell.value = formula
        else:
            cell.value = ''
        cell.alignment = cell_alignment
        cell.border = thin_border

        next_row += 1

    wb.save(MASTER_FILE)
    print(f'✓ Appended {len(triggers)} triggers to {MASTER_FILE}')


def load_master_results():
    """
    Load all results from Master_Results.xlsx and return as a DataFrame.
    Returns None if file doesn't exist or is empty.
    """
    if not os.path.exists(MASTER_FILE):
        return None

    try:
        with open(MASTER_FILE, 'rb') as f:
            df, _, _ = parse_results_upload(f.read())
        return None if df.empty else df
    except Exception as e:
        print(f'Warning: Could not load master results: {e}')
        return None


def get_cumulative_stats():
    """
    Calculate cumulative statistics from Master_Results.xlsx.
    Returns a dictionary with stats for each scenario.
    """
    df = load_master_results()
    if df is None or df.empty:
        return {}

    stats = {}
    
    # Group by scenario
    for scenario in df['Scenario'].unique():
        scenario_df = df[df['Scenario'] == scenario]
        
        # Count wins and losses
        wins = (scenario_df['Result'] == 'W').sum()
        losses = (scenario_df['Result'] == 'L').sum()
        total = wins + losses
        win_pct = wins / total if total > 0 else 0
        
        # Sum net P/L (only for W and L results)
        net_pl = scenario_df[scenario_df['Result'].isin(['W', 'L'])]['Net P/L'].sum()
        
        stats[scenario] = {
            'wins': wins,
            'losses': losses,
            'total': total,
            'win_pct': win_pct,
            'net_pl': net_pl
        }
    
    return stats


def get_master_results_path():
    """Return the full path to the master results file."""
    return os.path.abspath(MASTER_FILE)


if __name__ == '__main__':
    # Test: Initialize master results file
    initialize_master_results()
    print(f'\nMaster results file location: {get_master_results_path()}')
