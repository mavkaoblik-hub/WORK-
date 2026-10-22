"""Static v10.14 regression coverage (no Excel runtime or external downloads)."""
from pathlib import Path
import re
s=Path('deployment/demo-v10.14/TransportStatusFX.bas').read_text(encoding='utf-8-sig')
required = [
 'Private Sub ActiveFinancialSnapshot(', 'Private Sub ActiveRiskSignals(',
 'Private Function ArchiveMonthlyTable(', 'Private Sub ArchiveTrendChart(',
 'If S(trips.Cells(r, 8).Value2) = "У рейсі" Then',
 'If tripAge >= 14 Then', 'If forecast > 0 And cost > forecast Then',
 'ArchiveMonthlyTable(history, monthlyHeader, trips, byTrip)',
 'If lastMonthlyRow > 0 Then ArchiveTrendChart',
 'For i = sheet.ChartObjects.Count To 1 Step -1',
 'archiveFinanceEnd = categoriesEnd + 14',
 'AddMoney costs, monthId, Money(tripCosts, tripId)',
 'Дохід план USD',
 'Не оплачено USD', 'Витрати віднесено до місяця завершення рейсу',
]
for tag in required: assert tag in s, tag
assert s.count('Public Sub TC_RefreshBoth()') == 1
assert "Private Function Equivalent(" not in s, "Old FX helper can break compilation on Excel LTSC"
assert " = Equivalent(" not in s, "FinancialFooter must call the stable public FX converter"
assert s.count("Public Function TC_EquivalentUSD(") == 1
assert s.count("= TC_EquivalentUSD(") >= 3
assert "targetRate = Rate(targetISO)" in s
for kind in ['Sub','Function']:
    n1=len(re.findall(r'^\s*(?:(?:Public|Private|Friend)\s+)?'+kind+r'\s+\w+\s*\(',s,re.M))
    n2=len(re.findall(r'^\s*End '+kind+r'\s*$',s,re.M))
    assert n1==n2,(kind,n1,n2)
assert 'Invoke-WebRequest' not in s
p=Path('deployment/demo-v10.14/ASSEMBLE_PACKAGE.py').read_text(encoding='utf-8-sig')
for tag in ['demo-v10.13/TransportInsights.bas','demo-v10.14/TransportStatusFX.bas','demo-v10.14/INSTALL_DIAGNOSTIC.ps1','GENERATE_SEED.py','CREATE_DEMO_BASE.ps1']:
    assert tag in p,tag
print('PASS: v10.14 operational analytics and offline package dependency contract')
print('NOTE: Excel LTSC COM and VBA execution are NOT verified by this test.')

# v10.14.3: original error number/stage must survive; charts non-fatal.
assert s.count('refreshStage = "') == 27
assert 'sourceError = Err.Number' in s and 'sourceText = Err.Description' in s
assert 'TC_STATUS_FAILED' in s and 'errorText = "ЕТАП "' in s
inst=Path('deployment/demo-v10.14/INSTALL_DIAGNOSTIC.ps1').read_text(encoding='utf-8-sig')
assert '$refreshReady = $false' in inst
assert 'VBA_SMOKE_TEST_FAILED.txt' in inst
assert 'throw (\'VBA_REFRESH_FAILED:' in inst
assert p.find('INSTALL_DIAGNOSTIC.ps1') >= 0
print('PASS: v10.14.3 failure diagnostics, strict installer and package routing')

# v10.14.4: prevention of cmd.exe Unicode path corruption seen on Windows LTSC.
bootstrap=Path('deployment/demo-v10.14/BOOTSTRAP_INSTALL.ps1').read_bytes()
assert bootstrap.startswith(bytes((239, 187, 191))), 'PowerShell 5.1 needs Unicode BOM'
assert b'XLSM_EXISTS' in bootstrap and b'BOOTSTRAP_ERROR.txt' in bootstrap
assert 'BOOTSTRAP_INSTALL.ps1' in p
assert 'launcher.isascii()' in p
assert 'PowerShell' in p or 'powershell.exe' in p
print('PASS: ASCII-only launcher, BOM bootstrap and non-overwriting output')

# v10.14.5: guard FX display vs. USD ledger and run macros only on saved XLSM.
assert "On Error GoTo RateUnavailable" in s
assert "If fxSheet Is Nothing Then GoTo RateUnavailable" in s
assert "TC_FX_RATE_UNAVAILABLE:" in s
assert 'footerStep = fx & ": revenue"' in s
assert "TC_FINANCIAL_FOOTER_ERROR:" in s
source_inst=Path("deployment/demo-v10.14/INSTALL_DIAGNOSTIC.ps1").read_text(encoding="utf-8-sig")
save_pos=source_inst.index("$book.SaveAs($target,52)")
macro_pos=source_inst.index("[void]$excel.Run($runName)")
persist_pos=source_inst.index("$book.Save()")
assert save_pos < macro_pos < persist_pos, "VBA must run in saved XLSM"
assert source_inst.index("if (-not $refreshReady)") > persist_pos, "failure must be reported after saving diagnostics"
print("PASS: saved XLSM lifecycle and currency fallback/error stages")

# v10.14.6: readable UI, bounded Excel operations, single-pass expense ledger.
assert 'reportArea.Clear' in s and 'sheet.Cells.Clear' not in s
assert 'With reportArea.Font' in s and 'sheet.Cells.Font' not in s
assert 'detailRegion.Value2 = detailValues' in s
assert s.count('For er = ROW0 To Last(expenses, 2)') == 1
assert 'DisplayCategory(categoryText, id)' in s
assert 'lastMonthlyRow = ArchiveMonthlyTable(history, monthlyHeader, trips, byTrip)' in s
assert 'ActiveFinancialSnapshot live, Application.Max(20' in s
assert 'Application.ScreenUpdating = oldUpdating' in s
events=Path('deployment/demo-v10.14/TransportWorkbookEvents.txt').read_text(encoding='utf-8-sig')
assert 'mStatusLastRefresh As Date' in events
assert 'Workbook_SheetChange(ByVal Sh As Object' in events
assert 'DateValue(mStatusLastRefresh) = Date' in events
assert 'Private Sub RibbonRestore()' in events and 'Workbook_BeforeClose(Cancel As Boolean)' in events
print('PASS: v10.14.6 compact reports, batch receipts, fresh monthly insights, event-driven cache and ribbon restore')

# v10.14.7: COM input snapshot, manual update synchronization, no stale tab rebuild.
assert 'tripData = trips.Range("A" & ROW0 & ":AB" & lastTrip).Value' in s
assert 'expenseData = expenses.Range("A" & ROW0 & ":K" & lastExpense).Value' in s
assert 'tripData(tr - ROW0 + 1, 14)' in s
assert 'expenseData(er - ROW0 + 1, 8)' in s
assert 'ThisWorkbook.TC_DashboardRendered' in s
assert 'Public Sub TC_DashboardRendered()' in events
assert 'DateValue(mStatusLastRefresh) = Date' in events
assert 'DateDiff("s", mStatusLastRefresh, Now) < 20' not in events
assert 'Workbook_SheetChange(ByVal Sh As Object' in events
assert 'Application.EnableEvents = oldEvents' in s
print('PASS: v10.14.7 two ledger COM snapshots, change-driven refresh and manual sync')
