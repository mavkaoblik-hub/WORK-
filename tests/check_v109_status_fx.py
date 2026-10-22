"""Contract checks for the offline v10.9 administration dashboards and currency conversion."""
from pathlib import Path
import re

root=Path("deployment/offline-v10.9")
installer=(root/"INSTALL_OFFLINE.ps1").read_text(encoding="utf-8-sig")
app=(root/"TransportApp.bas").read_text(encoding="utf-8-sig")
fx=(root/"TransportStatusFX.bas").read_text(encoding="utf-8-sig")

for sheet in ("Поточний стан","Архів рейсів","Курси валют"):
    assert sheet in installer and sheet in fx, sheet
assert all(t in installer for t in (
    "$statusFxComponent.CodeModule.AddFromString($statusFxSource)",
    "DashboardButton $index 'Поточний стан'",
    "DashboardButton $index 'Архів рейсів'",
    "DashboardButton $index 'Курси валют'",
    "$book.SaveAs($target,52)",
    "$book.Close($false)",
))
assert installer.index("$book.SaveAs($target,52)") < installer.index("$book.Close($false)")
assert "codeload.github.com" not in installer
for proc in ("TC_RefreshBoth","TC_ShowLive","TC_ShowArchive","TC_ShowRates",
             "TC_RatesFetchNBP","TC_RatesManualApply","TC_USDFrom","TC_PLNPerUnit"):
    assert re.search(r"Public (?:Sub|Function) "+proc+r"\\(",fx),proc
assert "Equivalent(revenueUSD, fx)" in fx
assert "Equivalent(costUSD, fx)" in fx
assert 'Array("PLN", "USD", "EUR", "UAH")' in fx
assert '"немає курсу"' in fx
assert 'TripDays(planStart, actualEnd, True)' in fx
assert 'TripDays(planStart, actualEnd, False)' in fx
assert '"середня тривалість: "' in fx
assert 'stateOf.Exists(id)' in fx
assert fx.count("For tr = ROW0 To Last(trips, 1)") == 1, "N-squared expense scan"
assert 'CategorySummary(categoryTrip, id)' in fx
assert 'TC_RatesManualApply' in fx and 'TC_RatesFetchNBP' in fx
assert fx.count('req.Open "GET"') == 1
assert 'TransportStatusFX.TC_RefreshBoth' in app
assert 'TransportStatusFX.TC_USDFrom(amount, cur)' in app
assert 'TransportStatusFX.TC_PLNPerUnit(cur)' in app
assert '"EUR" And cur <> "UAH"' in app
assert 'TransportCenter.TC_RefreshNBP' not in app
print("v10.9: separate screens, indexed expenses, 4 FX equivalents, manual-only updating and builder order OK")
