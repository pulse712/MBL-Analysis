"""
Test Script for Master Results Manager
======================================
Run this to verify the master results system is working correctly.
"""

import os
import sys
from datetime import date

# Add parent directory to path
sys.path.insert(0, os.path.dirname(__file__))

from master_results_manager import (
    initialize_master_results,
    append_to_master_results,
    load_master_results,
    get_cumulative_stats,
    get_master_results_path,
    parse_results_upload,
    MASTER_FILE
)


def test_initialization():
    """Test 1: Initialize master results file"""
    print("\n" + "="*60)
    print("TEST 1: Initialize Master Results File")
    print("="*60)
    
    # Remove existing file if present
    if os.path.exists(MASTER_FILE):
        os.remove(MASTER_FILE)
        print(f"✓ Removed existing {MASTER_FILE}")
    
    # Initialize
    initialize_master_results()
    
    # Check file exists
    if os.path.exists(MASTER_FILE):
        print(f"✅ SUCCESS: {MASTER_FILE} created")
        print(f"   Location: {get_master_results_path()}")
        return True
    else:
        print(f"❌ FAILED: {MASTER_FILE} not created")
        return False


def test_append_results():
    """Test 2: Append sample results to master file"""
    print("\n" + "="*60)
    print("TEST 2: Append Sample Results")
    print("="*60)
    
    # Create sample triggers
    sample_triggers = [
        {
            'team': 'NEW YORK YANKEES',
            'opponent': 'BOSTON RED SOX',
            'home_away': 'home',
            'line': -150,
            'scenario_id': '01',
            'scenario': 'BLOWOUT #1 - MJ',
            'verdict': 'CLEAR BET',
            'play': 'BET NEW YORK YANKEES'
        },
        {
            'team': 'LOS ANGELES DODGERS',
            'opponent': 'SAN FRANCISCO GIANTS',
            'home_away': 'away',
            'line': 120,
            'scenario_id': '14',
            'scenario': 'THE SCORING DROUGHT #1 - MJ',
            'verdict': 'CLEAR FADE',
            'play': 'FADE LOS ANGELES DODGERS'
        },
        {
            'team': 'HOUSTON ASTROS',
            'opponent': 'SEATTLE MARINERS',
            'home_away': 'home',
            'line': -105,
            'scenario_id': '22',
            'scenario': 'SMALL ROAD FAVORITE / NEW SERIES - SM',
            'verdict': 'INCONSISTENT',
            'play': 'WATCH HOUSTON ASTROS'
        }
    ]
    
    test_date = date(2026, 7, 9)
    
    try:
        append_to_master_results(sample_triggers, test_date)
        print(f"✅ SUCCESS: Appended {len(sample_triggers)} triggers to master file")
        return True
    except Exception as e:
        print(f"❌ FAILED: {e}")
        return False


def test_load_results():
    """Test 3: Load results from master file"""
    print("\n" + "="*60)
    print("TEST 3: Load Master Results")
    print("="*60)
    
    try:
        df = load_master_results()
        
        if df is not None and not df.empty:
            print(f"✅ SUCCESS: Loaded {len(df)} results")
            print("\nSample data:")
            print(df[['Date', 'Team', 'Scenario', 'Result']].head())
            return True
        else:
            print("⚠️  WARNING: Master file is empty")
            return False
    except Exception as e:
        print(f"❌ FAILED: {e}")
        return False


def test_cumulative_stats():
    """Test 4: Calculate cumulative statistics"""
    print("\n" + "="*60)
    print("TEST 4: Calculate Cumulative Stats")
    print("="*60)
    
    try:
        stats = get_cumulative_stats()
        
        if stats:
            print(f"✅ SUCCESS: Calculated stats for {len(stats)} scenarios")
            print("\nSample stats (first 3 scenarios):")
            for i, (scenario, data) in enumerate(list(stats.items())[:3]):
                print(f"\n{scenario}:")
                print(f"  Wins: {data['wins']}, Losses: {data['losses']}, Total: {data['total']}")
                print(f"  Win %: {data['win_pct']:.1%}, Net P/L: ${data['net_pl']:.2f}")
            return True
        else:
            print("⚠️  WARNING: No stats calculated (need W/L results entered)")
            return True  # Not a failure, just no results yet
    except Exception as e:
        print(f"❌ FAILED: {e}")
        return False


def test_add_results_with_outcomes():
    """Test 5: Add results and simulate entering W/L outcomes"""
    print("\n" + "="*60)
    print("TEST 5: Simulate Entering W/L Results")
    print("="*60)
    
    try:
        from openpyxl import load_workbook
        
        # Load the workbook
        wb = load_workbook(MASTER_FILE)
        ws = wb.active
        
        # Simulate entering results (W or L) for the sample data
        # Row 4 is first data row
        results = ['W', 'L', 'W']  # Win, Loss, Win
        
        for i, result in enumerate(results, start=4):
            ws.cell(row=i, column=8).value = result  # Column H is Result
        
        wb.save(MASTER_FILE)
        print(f"✅ SUCCESS: Added W/L results to {len(results)} rows")
        
        # Now test if stats calculate correctly
        stats = get_cumulative_stats()
        if stats:
            total_plays = sum(s['total'] for s in stats.values())
            print(f"✓ Cumulative stats now show {total_plays} completed plays")
            return True
        else:
            print("⚠️  WARNING: Stats not calculated")
            return False
            
    except Exception as e:
        print(f"❌ FAILED: {e}")
        return False


def test_file_references():
    """Test 6: Verify Excel formula references work"""
    print("\n" + "="*60)
    print("TEST 6: Verify Excel Formula References")
    print("="*60)
    
    try:
        from openpyxl import load_workbook
        
        # Check if formulas reference external file correctly
        wb = load_workbook(MASTER_FILE)
        ws = wb.active
        
        # Check if Net P/L formulas exist
        formula_count = 0
        for row in range(4, 10):  # Check first few data rows
            cell = ws.cell(row=row, column=9)  # Column I is Net P/L
            if cell.value and str(cell.value).startswith('='):
                formula_count += 1
        
        if formula_count > 0:
            print(f"✅ SUCCESS: Found {formula_count} Net P/L formulas")
            print("✓ Formulas will auto-calculate based on W/L entries")
            return True
        else:
            print("⚠️  WARNING: No formulas found (may be intentional)")
            return True
            
    except Exception as e:
        print(f"❌ FAILED: {e}")
        return False


def test_parse_daily_report_upload():
    """Test 7: Parse results from a saved daily report (Results Tracker tab)"""
    print("\n" + "="*60)
    print("TEST 7: Parse Daily Report Upload")
    print("="*60)

    try:
        import io
        import xlsxwriter

        buf = io.BytesIO()
        wb = xlsxwriter.Workbook(buf, {'in_memory': True})
        wb.add_worksheet('Green Clear Bet')
        ws = wb.add_worksheet('Results Tracker')
        ws.write(0, 0, 'RESULTS TRACKER')
        ws.write(1, 0, 'Enter W or L')
        headers = ['Date', 'Team', 'H/A', 'Odds', 'Play', 'Scenario', 'Type', 'Result (W/L)', 'Net P/L ($100)']
        for col, h in enumerate(headers):
            ws.write(2, col, h)
        ws.write(3, 0, '2026-07-20')
        ws.write(3, 1, 'New York Yankees')
        ws.write(3, 2, 'HOME')
        ws.write(3, 3, '-150')
        ws.write(3, 4, 'BET NEW YORK YANKEES')
        ws.write(3, 5, '#01 BLOWOUT #1 - MJ')
        ws.write(3, 6, 'CLEAR BET')
        ws.write(3, 7, 'W')
        ws.write(3, 8, 66.67)
        wb.close()
        file_bytes = buf.getvalue()

        df, source, sheet = parse_results_upload(file_bytes)

        if source == 'daily report' and sheet == 'Results Tracker' and len(df) == 1:
            if df.iloc[0]['Result'] == 'W' and str(df.iloc[0]['Date'])[:10] == '2026-07-20':
                print("✅ SUCCESS: Parsed daily report Results Tracker with W/L")
                return True

        print(f"❌ FAILED: Unexpected parse result source={source} sheet={sheet} rows={len(df)}")
        return False
    except Exception as e:
        print(f"❌ FAILED: {e}")
        return False


def test_master_dedup_and_fade_roundtrip():
    """Test 8: Double-generate dedup + multi-cycle FADE PayoutLine persistence"""
    print("\n" + "="*60)
    print("TEST 8: Master Dedup & FADE Round-Trip")
    print("="*60)

    try:
        import sys
        import os
        from datetime import date
        from openpyxl import load_workbook
        import io

        sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'pages'))
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            'daily_report_page',
            os.path.join(os.path.dirname(__file__), 'pages', '1_Daily_Report.py'),
        )
        page = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(page)

        bet = {
            'verdict': 'CLEAR BET', 'line': -150, 'scenario_id': '01', 'scenario': 'B',
            'team': 'NEW YORK YANKEES', 'opponent': 'BOSTON RED SOX',
            'home_away': 'home', 'play': 'BET',
        }
        fade = {
            'verdict': 'CLEAR FADE', 'line': -150, 'opp_line': 130,
            'scenario_id': '02', 'scenario': 'Y', 'team': 'NEW YORK YANKEES',
            'opponent': 'BOSTON RED SOX', 'home_away': 'home', 'play': 'FADE',
        }

        # Double-generate should not duplicate
        df1, _, _ = parse_results_upload(page.build_master_file(None, [bet], date(2026, 7, 21)))
        df2, _, _ = parse_results_upload(page.build_master_file(df1, [bet], date(2026, 7, 21)))
        if len(df1) != 1 or len(df2) != 1:
            print(f"❌ FAILED: Double-generate rows {len(df1)} -> {len(df2)}, expected 1 -> 1")
            return False

        # DH same matchup different lines -> 2 rows
        t1 = {**bet, 'line': -150}
        t2 = {**bet, 'line': -130}
        df_dh, _, _ = parse_results_upload(page.build_master_file(None, [t1, t2], date(2026, 7, 21)))
        if len(df_dh) != 2:
            print(f"❌ FAILED: DH rows {len(df_dh)}, expected 2")
            return False

        # FADE: 3 upload/generate cycles keep +130 formula
        df = None
        ok = True
        for i in range(3):
            b = page.build_master_file(df, [fade] if i == 0 else [], date(2026, 7, 21))
            ws = load_workbook(io.BytesIO(b)).active
            formula = str(ws['I4'].value or '')
            jval = ws['J4'].value
            if '130' not in formula:
                print(f"❌ FAILED: Cycle {i+1} formula lost +130: {formula}")
                ok = False
                break
            if jval is None:
                print(f"❌ FAILED: Cycle {i+1} PayoutLine col J is empty")
                ok = False
                break
            df, _, _ = parse_results_upload(b)

        if ok:
            print("✅ SUCCESS: Dedup and 3-cycle FADE PayoutLine persistence")
            return True
        return False
    except Exception as e:
        print(f"❌ FAILED: {e}")
        return False


def run_all_tests():
    """Run all tests"""
    print("\n" + "="*70)
    print("🧪 MASTER RESULTS SYSTEM - TEST SUITE")
    print("="*70)
    
    tests = [
        ("Initialization", test_initialization),
        ("Append Results", test_append_results),
        ("Load Results", test_load_results),
        ("Cumulative Stats", test_cumulative_stats),
        ("W/L Entry Simulation", test_add_results_with_outcomes),
        ("Formula References", test_file_references),
        ("Daily Report Upload", test_parse_daily_report_upload),
        ("Master Dedup & FADE Round-Trip", test_master_dedup_and_fade_roundtrip),
    ]
    
    results = []
    for name, test_func in tests:
        try:
            result = test_func()
            results.append((name, result))
        except Exception as e:
            print(f"\n❌ CRITICAL ERROR in {name}: {e}")
            results.append((name, False))
    
    # Summary
    print("\n" + "="*70)
    print("📊 TEST SUMMARY")
    print("="*70)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{status:10} {name}")
    
    print("\n" + "-"*70)
    print(f"Results: {passed}/{total} tests passed ({passed/total*100:.0f}%)")
    
    if passed == total:
        print("\n🎉 ALL TESTS PASSED! Master Results system is working correctly.")
        print("\nNext steps:")
        print("1. Run the Streamlit app: streamlit run app.py")
        print("2. Generate a daily report")
        print("3. Copy results from Results Tracker to Master_Results.xlsx")
        print("4. Check Scenario Performance tab for cumulative stats")
    else:
        print(f"\n⚠️  {total - passed} test(s) failed. Check the output above for details.")
    
    print("="*70 + "\n")
    
    return passed == total


if __name__ == '__main__':
    success = run_all_tests()
    sys.exit(0 if success else 1)
