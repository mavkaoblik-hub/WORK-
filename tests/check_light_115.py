"""v10.15 LIGHT dashboard source checks: no Excel runtime needed."""
from pathlib import Path
import re

root=Path("deployment/light-v10.15")
theme=(root/"DESIGN_LIGHT.ps1").read_text(encoding="utf-8-sig")
installer=(root/"INSTALL_LIGHT.ps1").read_text(encoding="utf-8-sig")
status=(root/"TransportStatusFX.bas").read_text(encoding="utf-8-sig")
base=(Path("deployment/clean-v10.14.7/INSTALL_CLEAN.ps1")).read_text(encoding="utf-8-sig")
assert "DESIGN_LIGHT.ps1" in installer
assert "Join-Path $dir 'DESIGN_LIGHT.ps1'" in installer
assert "-Sheet $index" in installer
assert "$ws.Range('A1:R80').UnMerge()" in theme
assert "RGB 245 248 250" in theme and "RGB 255 255 255" in theme
assert "Segoe UI" in theme
assert theme.count(" @('") >= 8, "Eight KPI cards required"
for label in ("ВОДІЇ", "ЗАПЛАНОВАНО", "У РЕЙСІ", "ЗАВЕРШЕНО", "ВИТРАТИ", "ДОХІД"):
    assert label in theme
assert "('~',[string][char]39)" in theme or ".Replace('~',[string][char]39)" in theme
for label in (
    "Створити рейс","Розпочати рейс","Завершити рейс","Внести витрати",
    "Автопарк","Водії / контакти","Маршрути Європи","Паркінги",
    "Фінансовий звіт","Резервна копія","Оновити NBP вручну","Поточний стан",
    "Оновити пробіг","Архів рейсів","Курси валют","Відновити копію"
):
    assert installer.count(f"DashboardButton $index '{label}'") == 1, label
assert "DEMO_SEED.ps1" not in installer and "demo_seed.json" not in installer
assert "CLEAN DATA CHECK" in installer and "$book.SaveAs($target,52)" in installer
assert "ByVal currency As String" not in status
assert "TC_USDFrom(ByVal amount As Double, ByVal currencyCode As String)" in status
assert "RGB(238, 248, 244)" in status
assert "RGB(233, 246, 240)" in status
assert "RGB(12, 36, 31)" in base  or "forest-themed dashboard" in base
assert "forest-themed dashboard" not in installer
assert "TC_ShowArchive" in installer and "TransportInsights.TC_GoHome" in installer
for kind in ("Sub","Function"):
    n1=len(re.findall(r"^\s*(?:Public|Private)\s+"+kind+r"\s+\w+\s*\(",status,re.M))
    n2=len(re.findall(r"^\s*End "+kind+r"\s*$",status,re.M))
    assert n1==n2,(kind,n1,n2)
print("PASS: v10.15 light color/style, 8 KPI cards, 16 intact actions, blank ledger, FX compatibility")
print("NOTE: Windows Excel LTSC/VBA execution is not exercised.")

# v10.15.1 - address the user-observed small typography and archive opening at row 16.
events = Path("deployment/demo-v10.14/TransportWorkbookEvents.txt").read_text(encoding="utf-8-sig")
assert "Private mArchiveFirstView As Boolean" in events
assert "Private mLiveFirstView As Boolean" in events
assert "mArchiveFirstView = True" in events
assert "mLiveFirstView = True" in events
assert "ActiveWindow.ScrollRow = 1" in events
assert "ActiveWindow.ScrollColumn = 1" in events
assert "ActiveWindow.Zoom = 92" in events
assert "mArchiveFirstView = False" in events
assert "mLiveFirstView = False" in events
assert "$ws.Range([string]$item[5]).Font.Size=21" in theme
assert "$ws.Range([string]$item[3]).Font.Size=11" in theme
assert "$ws.Range('B2').Font.Size=23" in theme
assert "ActiveWindow.Zoom = 92" in status
assert "$excel.ActiveWindow.Zoom=92" in installer
assert "$s.TextFrame.Characters().Font.Size = 12" in installer
assert "DEMO_SEED.ps1" not in installer
print("PASS: v10.15.1 first archive visit shows title; subsequent visits preserve user scroll")
