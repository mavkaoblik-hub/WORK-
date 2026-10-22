"""Static v10.15.2 LIGHT regression: navigation, rate inputs and readable empty-state."""
from pathlib import Path
import re

base = Path("deployment/light-v10.15")
theme = (base / "DESIGN_LIGHT.ps1").read_text(encoding="utf-8-sig")
install = (base / "INSTALL_LIGHT.ps1").read_text(encoding="utf-8-sig")
fx = (base / "TransportStatusFX.bas").read_text(encoding="utf-8-sig")
events = Path("deployment/demo-v10.14/TransportWorkbookEvents.txt").read_text(encoding="utf-8-sig")

for target in ("L4:M4", "N4:O4", "P4:Q4"):
    assert target in theme, target
for anchor in ("#''Зміст''!B15", "#''Зміст''!B20", "#''Зміст''!B25"):
    assert anchor in theme, anchor
assert "$ws.Range('B4:K4').Merge()" in theme
assert "$ws.Range('B4:Q4').Merge()" not in theme
assert theme.count("=HYPERLINK(") == 3
assert "CLEAN · v10.15.2" in theme

for required in (
    "$rates.Range('B7:B9').ClearContents()",
    "$rates.Range('B7:B9').Interior.Color",
    "$rates.Range('H5:M5').Merge()",
    "$rates.Range('H6:M11').Merge()",
    "DashboardButton $rates 'Оновити NBP' 'TC_RatesFetchNBP' 'A14'",
    "DashboardButton $rates 'Застосувати ручні курси' 'TC_RatesManualApply' 'D14'",
    "DashboardButton $rates 'На головну' 'TC_GoHome' 'A18'",
):
    assert required in install, required
assert "DEMO_SEED.ps1" not in install and "demo_seed.json" not in install
assert "CLEAN DATA CHECK" in install
assert install.count("DashboardButton $index '") == 16

for required in (
    "With sh.Range(sh.Cells(atRow + 1, 5), sh.Cells(atRow + 1, 12))",
    "With sh.Range(sh.Cells(atRow + 2, 5), sh.Cells(atRow + 2, 12))",
    "Планова виручка мінус зареєстровані витрати",
    "If persons.Count = 0 And completed = 0 And active = 0 And planned = 0 Then",
    'With live.Range("A13:L14")',
    'With history.Range("A13:L16")',
    "ActiveFinancialSnapshot live, IIf(persons.Count = 0, 15,",
):
    assert required in fx, required
assert "ByVal currency As String" not in fx
assert "mArchiveFirstView" in events and "ActiveWindow.ScrollRow = 1" in events
assert len(re.findall(r"^\s*(?:Public|Private)\s+Sub\s+\w+\s*\(", fx, re.M)) == len(re.findall(r"^\s*End Sub\s*$", fx, re.M))
assert len(re.findall(r"^\s*(?:Public|Private)\s+Function\s+\w+\s*\(", fx, re.M)) == len(re.findall(r"^\s*End Function\s*$", fx, re.M))
print("PASS v10.15.2: light UI nav, safe FX inputs, 16 macro actions and no fixture data.")
print("NOTE: Excel LTSC VBA execution and live hyperlink activation require manual validation.")
