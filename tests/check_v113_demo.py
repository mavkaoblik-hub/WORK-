"""GitHub-source static regression check; full seeded demo QA runs inside customer ZIP."""
from pathlib import Path
import re
s=Path('deployment/demo-v10.13/TransportInsights.bas').read_text(encoding='utf-8-sig')
required=[
 'Private Function AppendExpenseLedger(',
 'Private Function AppendMonthlySummary(',
 'Private Function EarliestActualMonth(',
 'endRow = AppendMonthlySummary(w, endRow + 3',
 'endRow = AppendExpenseLedger(w, endRow + 3',
 'w.PageSetup.PrintArea = "$A$1:$L$" & endRow',
 'w.PageSetup.PrintTitleRows = "$1:$3"',
 'If driverFilter = "" Or driverFilter = driverId Then',
 'Within(whenSpent, firstDay, lastDay)',
 'TransportStatusFX.TC_EquivalentUSD',
 'ДЕТАЛІЗАЦІЯ ЗАВЕРШЕНИХ РЕЙСІВ',
 'КНИГА ВИТРАТ',
 'ПОМІСЯЧНА ДИНАМІКА',
 'За весь період станом на',
 'Усі водії (компанія)',
]
for tag in required:
 assert tag in s,tag
assert 'If reportMode = "ALL" Then monthCount = 12' not in s
assert s.count('AddChart w,')==3
for typ in ('Sub','Function'):
 starts=len(re.findall(r'^\s*(?:(?:Public|Private|Friend)\s+)?'+typ+r'\s+\w+\s*\(',s,re.M))
 ends=len(re.findall(r'^\s*End '+typ+r'\s*$',s,re.M))
 assert starts==ends,(typ,starts,ends)
print('PASS: monthly financial table, complete receipts ledger, 3 charts and all/selected-driver filtering')
print('Runtime VBA is not exercised in GitHub hosted runner.')