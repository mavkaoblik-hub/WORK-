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
for tag in ['demo-v10.13/TransportInsights.bas','demo-v10.14/TransportStatusFX.bas','demo-v10.12/INSTALL_OFFLINE.ps1','GENERATE_SEED.py','CREATE_DEMO_BASE.ps1']:
    assert tag in p,tag
print('PASS: v10.14 operational analytics and offline package dependency contract')
print('NOTE: Excel LTSC COM and VBA execution are NOT verified by this test.')
