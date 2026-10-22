"""Static regression for transport finance report: all drivers by default, or one selected driver."""
from pathlib import Path
import re

root = Path("deployment/offline-v10.8")
app = (root / "TransportApp.bas").read_text(encoding="utf-8-sig")
report = (root / "TransportInsights.bas").read_text(encoding="utf-8-sig")
installer = (root / "INSTALL_OFFLINE.ps1").read_text(encoding="utf-8-sig")

assert 'If mMode = "REPORT" Then f.Controls("cboDriver").AddItem "Усі водії (компанія)"' in app
assert 'f.Controls("cboDriver").ListIndex = 0' in app
assert 'EnableCombo f, "Driver", "Для кого звіт"' in app
assert 'driverText <> "Усі водії (компанія)"' in app
assert 'driverId = DriverIdFromDisplay(driverText)' in app
assert 'TransportInsights.TC_GenerateReport reportMode, startDate, endDate, driverId' in app
assert 'Case "REPORT": SaveReport f' in app
assert 'For Each entry In mUiDrivers: f.Controls("cboDriver").AddItem CStr(entry): Next entry' in app

# The optional blank driverFilter signifies ALL; populated filter isolates one driver
# for every relevant metric and chart source.
assert 'Optional ByVal driverFilter As String = ""' in report
assert report.count('If Len(driverFilter) = 0 Or driverFilter = driverId Then') >= 2
assert report.count('If Len(driverFilter) = 0 Or driverFilter = id Then') >= 2
assert 'If Len(driverFilter) = 0 Then' in report
assert 'labelPeriod = labelPeriod & " | Усі водії (компанія)"' in report
assert '"ПОКАЗНИКИ ВИБРАНОГО ВОДІЯ' in report
assert '"ПОКАЗНИКИ ЗА ВСІМА ВОДІЯМИ' in report
assert 'Transport_Report_ALL_' in report
assert 'Transport_Report_DRIVER_' in report
assert 'w.ExportAsFixedFormat Type:=xlTypePDF' in report
assert report.count('AddChart w,') == 3
assert 'WS("Зміст").ExportAsFixedFormat' not in app

# Both reporting and current-status VBA are injected into workbook, offline.
assert '$insightsFile' in installer
assert "$insightsComponent.Name='TransportInsights'" in installer
assert "'Поточний стан'" in installer
assert "$book.SaveAs($target,52)" in installer
assert 'codeload.github.com' not in installer
print("v10.8 driver-scope report and current-status static contract: OK")
