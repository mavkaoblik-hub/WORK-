from pathlib import Path
import json,re,zipfile
from decimal import Decimal
from collections import Counter,defaultdict
p=Path(__file__).resolve().parent
vba=(p/'vba/TransportStatusFX.bas').read_text(encoding='utf-8-sig')
rep=(p/'vba/TransportInsights.bas').read_text(encoding='utf-8-sig')
events=(p/'vba/TransportWorkbookEvents.txt').read_text(encoding='utf-8-sig')
installer=(p/'INSTALL_OFFLINE.ps1').read_text(encoding='utf-8-sig')
seed=json.loads((p/'data/demo_seed.json').read_text(encoding='utf-8-sig'))
expected=json.loads((p/'data/demo_expected.json').read_text(encoding='utf-8-sig'))
trips=seed['trips']; exp=seed['expenses']
assert len(seed['drivers'])==7 and len(seed['vehicles'])==9
assert len(trips)==75 and len(exp)==259
assert Counter(x['status'] for x in trips)=={'Завершено':70,'У рейсі':3,'Заплановано':2}
roundD=lambda x: Decimal(str(x))
completed={t['id'] for t in trips if t['status']=='Завершено'}
assert sum((roundD(t['revenue_usd']) for t in trips if t['id'] in completed),Decimal(0))==roundD(expected['completed_income_usd'])
assert sum((roundD(e['usd']) for e in exp),Decimal(0))==roundD(expected['all_expenses_usd'])
assert sum((roundD(e['usd']) for e in exp if e['trip'] in completed),Decimal(0))==roundD(expected['completed_expenses_usd'])
# Trip IDs, references, driver list and expense date parse cleanly.
ids=[x['id'] for x in trips]; assert len(ids)==len(set(ids))
assert {x['driver'] for x in trips}.issubset({d['id'] for d in seed['drivers']})
assert {x['car'] for x in trips}.issubset({d['id'] for d in seed['vehicles']})
assert {x['trip'] for x in exp}.issubset(set(ids))
# Financial breakdown must avoid double counted amounts and use selected driver scope.
assert 'Private Function ReportAppendices(' in rep
assert 'TransportStatusFX.TC_EquivalentUSD' in rep
assert 'endRow = ReportAppendices(' in rep
assert 'w.PageSetup.PrintArea = "$A$1:$L$" & endRow' in rep
assert 'РОЗШИФРУВАННЯ ВИТРАТ ЗА КАТЕГОРІЯМИ' in rep
assert 'ЕФЕКТИВНІСТЬ АВТОПАРКУ' in rep
assert 'ДЕТАЛІЗАЦІЯ ЗАВЕРШЕНИХ РЕЙСІВ' in rep
assert 'If (driverFilter = "" Or driverFilter = driver) And Within(dayCost, firstDay, lastDay)' in rep
assert 'Within(dayFinished, firstDay, lastDay)' in rep
assert rep.count('AddChart w,')==3
assert 'Public Function TC_EquivalentUSD(' in vba
for word in ('DriverOverview live,','DriverOverview history,','byCategoryAll','byCategoryFinished', 'Array("PLN", "USD", "EUR", "UAH")'):
 assert word in vba,word
assert events.count('Private Sub Workbook_Activate()')==1
assert events.count('Private Sub Workbook_Deactivate()')==1
assert events.count('Private Sub Workbook_WindowDeactivate(')==1
assert events.count('Private Sub Workbook_BeforeClose(')==1
assert 'RibbonRestore' in events and 'RibbonForTransport' in events
assert 'Application.CommandBars("Ribbon").Visible' in events
assert not any(x in events for x in ('req.Open','NbpRate','TC_RatesFetchNBP','Invoke-WebRequest'))
assert '$workbookComponent.CodeModule.AddFromString($eventSource)' in installer
assert installer.index('$book.SaveAs($target,52)')<installer.index('$book.Close($false)')
assert 'codeload.github.com' not in installer
for name in ('TransportApp.bas','TransportInsights.bas','TransportStatusFX.bas','TransportCenter.bas','TransportWorkbookEvents.txt'):
 text=(p/'vba'/name).read_text(encoding='utf-8-sig')
 for kind in ('Sub','Function'):
  starts=len(re.findall(r'^\s*(?:(?:Public|Private|Friend)\s+)?'+kind+r'\s+\w+\s*\(',text,re.M))
  ends=len(re.findall(r'^\s*End '+kind+r'\s*$',text,re.M))
  assert starts==ends,(name,kind,starts,ends)
 odd=[(i,x) for i,x in enumerate(text.splitlines(),1) if not x.strip().startswith("'") and re.sub(r'""','',x).count('"')%2]
 assert not odd,(name,odd[:2])
with zipfile.ZipFile(p/'Транспортний_центр_v8_робочий_Times_New_Roman.xlsx') as z:
 assert z.testzip() is None
 assert len([x for x in z.namelist() if '/charts/chart' in x and x.endswith('.xml')])==4
print('PASS: 75 demo trips, 259 expenses, 7 drivers, accounting sums, 3 report charts,')
print('      dynamic driver dashboards, selected-driver details, 4 currencies, ribbon restore, offline installer.')
print('Excel LTSC VBA runtime is NOT tested by this static validation.')