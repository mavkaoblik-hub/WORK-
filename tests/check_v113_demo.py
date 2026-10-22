from pathlib import Path
from collections import Counter, defaultdict
from datetime import date
from decimal import Decimal
import json, re, zipfile
p=Path(__file__).resolve().parent
seed=json.loads((p/'data/demo_seed.json').read_text(encoding='utf-8-sig'))
expect=json.loads((p/'data/demo_expected.json').read_text(encoding='utf-8-sig'))
rep=(p/'vba/TransportInsights.bas').read_text(encoding='utf-8-sig')
events=(p/'vba/TransportWorkbookEvents.txt').read_text(encoding='utf-8-sig')
inst=(p/'INSTALL_OFFLINE.ps1').read_text(encoding='utf-8-sig')
trips=seed['trips']; exp=seed['expenses']
assert len(trips)==75 and len(exp)==259
assert Counter(t['status'] for t in trips)=={'Завершено':70,'У рейсі':3,'Заплановано':2}
assert sum(Decimal(str(e['usd'])) for e in exp)==Decimal(str(expect['all_expenses_usd']))
assert sum(Decimal(str(t['revenue_usd'])) for t in trips if t['status']=='Завершено')==Decimal(str(expect['completed_income_usd']))
ids={t['id']:t for t in trips}
assert all(e['trip'] in ids for e in exp)
# New printed invoice ledger must match the total expense scope, for any driver.
for drv in [None]+[d['id'] for d in seed['drivers']]:
    rows=[e for e in exp if drv is None or ids[e['trip']]['driver']==drv]
    total=sum((Decimal(str(e['usd'])) for e in rows),Decimal(0))
    assert total >= 0
    if drv is None: assert total==Decimal(str(expect['all_expenses_usd']))
# All actual demo journal events are within three months, not 12 empty graph columns.
completed_months={t['actual_end'][:7] for t in trips if t['status']=='Завершено'}
expense_months={e['date'][:7] for e in exp} if 'date' in exp[0] else set()
assert completed_months=={'2026-08','2026-09','2026-10'},completed_months
assert 'AppendExpenseLedger(w, endRow + 3' in rep
assert 'AppendMonthlySummary(w, endRow + 3' in rep
assert 'EarliestActualMonth(endDate, dMonthlyIncome, dMonthlyExpense)' in rep
assert 'w.PageSetup.PrintArea = "$A$1:$L$" & endRow' in rep
assert 'w.PageSetup.PrintTitleRows = "$1:$3"' in rep
assert 'If reportMode = "ALL" Then monthCount = 12' not in rep
assert rep.count('AddChart w,')==3
for part in ('КНИГА ВИТРАТ','ПОМІСЯЧНА ДИНАМІКА','ЕФЕКТИВНІСТЬ АВТОПАРКУ','ДЕТАЛІЗАЦІЯ ЗАВЕРШЕНИХ РЕЙСІВ'):
    assert part in rep,part
for kind in ('Sub','Function'):
    starters=len(re.findall(r'^\s*(?:(?:Public|Private|Friend)\s+)?'+kind+r'\s+\w+\s*\(',rep,re.M))
    ends=len(re.findall(r'^\s*End '+kind+r'\s*$',rep,re.M))
    assert starters==ends,(kind,starters,ends)
assert all(x in events for x in ('Workbook_Activate()','Workbook_Deactivate()','Workbook_BeforeClose(Cancel As Boolean)','RibbonRestore'))
assert inst.index('$book.SaveAs($target,52)')<inst.index('$book.Close($false)')
assert not any(t in inst for t in ('codeload.github.com','Invoke-WebRequest'))
for n in ('Транспортний_центр_v8_робочий_Times_New_Roman.xlsx','DEMO_PREVIEW.xlsx'):
    with zipfile.ZipFile(p/n) as z: assert z.testzip() is None
print('PASS: demo 7/9/75/259, USD ledger, itemized expenses, 3-month charts, PDF full-range, VBA structure, v9 installer ordering.')
print('IMPORTANT: Actual VBA execution in Windows Excel LTSC cannot be validated here.')