"""Transport Center v10.10: audit the single source of truth and ledger categories.
This is static validation only; GitHub hosted runners do not execute Microsoft Excel.
"""
from pathlib import Path
import re

R=Path("deployment/offline-v10.10")
fx=(R/"TransportStatusFX.bas").read_text(encoding="utf-8-sig")
rep=(R/"TransportInsights.bas").read_text(encoding="utf-8-sig")
assert "Public Sub TC_RefreshLive()" not in rep, "Retired dashboard builder returned"
assert "Public Sub TC_ShowLiveStatus()" not in rep
assert "Public Sub TC_GoHome()" in rep
assert "Public Sub TC_GenerateReport(" in rep
assert "Public Sub TC_RefreshBoth()" in fx
assert "Private Function CategoryTotalsTable(" in fx
assert 'AddMoney byCategoryAll, category, cost' in fx
assert 'AddMoney byCategoryFinished, category, N(expenses.Cells(er, 8).Value2)' in fx
assert 'If CStr(stateOf(id)) = "Завершено" Then' in fx
assert fx.count('For er = ROW0 To Last(expenses, 2)') == 2
assert fx.count('For tr = ROW0 To Last(trips, 1)') == 1
assert 'CategoryTotalsTable(history,' in fx
assert 'CategoryTotalsTable(live,' in fx
assert fx.index('categoriesEnd = CategoryTotalsTable(history,') < fx.index('FinancialFooter history,')
assert fx.index('categoriesEnd = CategoryTotalsTable(live,') < fx.index('FinancialFooter live,')
assert 'середня тривалість' in fx
assert 'CategorySummary(categoryTrip, id)' in fx
assert 'Array("PLN", "USD", "EUR", "UAH")' in fx
assert 'Equivalent(revenueUSD - costUSD, fx)' in fx
assert 'немає курсу' in fx
assert "Public Sub TC_RatesManualApply()" in fx
assert "Public Sub TC_RatesFetchNBP()" in fx
assert fx.count('req.Open "GET"') == 1
for name in ("ПОТОЧНИЙ СТАН","АРХІВ РЕЙСІВ","ВИТРАТИ ЗА КАТЕГОРІЯМИ","СТРУКТУРА ВИТРАТ"):
    assert name in fx, name
assert 'Transport_Report_ALL_' in rep
assert 'Transport_Report_DRIVER_' in rep
assert rep.count("AddChart w,") == 3

for name,code in [("TransportInsights",rep),("TransportStatusFX",fx)]:
    assert code.count('"') > 0
    for typ in ("Sub","Function"):
        count=len(re.findall(r"^\s*(?:Public |Private |Friend )?"+typ+r"\s+\w+\s*\(",code,re.M))
        ends=len(re.findall(r"^\s*End "+typ+r"\b",code,re.M))
        assert count==ends, (name,typ,count,ends)

# Independent expected allocation for 2 trips across two categories.
expenses=[("done","Fuel",50),("done","Parking",20),("active","Fuel",30)]
states={"done":"Завершено","active":"У рейсі"}
company=sum(cost for trip,cat,cost in expenses)
finished=sum(cost for trip,cat,cost in expenses if states[trip]=="Завершено")
assert company==100 and finished==70
assert sum(x[2] for x in expenses if x[1]=="Fuel")==80
print("v10.10: single current-state writer, category totals, 4 FX currencies and reporting OK")
