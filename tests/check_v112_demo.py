"""Static v10.12 review on GitHub runner; customer ZIP carries full demo-seed financial checks."""
from pathlib import Path
import re
r=Path("deployment/demo-v10.12")
inst=(r/"INSTALL_OFFLINE.ps1").read_text(encoding="utf-8-sig")
rep=(r/"TransportInsights.bas").read_text(encoding="utf-8-sig")
fx=(r/"TransportStatusFX.bas").read_text(encoding="utf-8-sig")
ev=(r/"TransportWorkbookEvents.txt").read_text(encoding="utf-8-sig")
assert "Private Function ReportAppendices(" in rep
assert "endRow = ReportAppendices(" in rep
assert "w.PageSetup.PrintArea = " in rep
assert rep.count("AddChart w,") == 3
assert "If (driverFilter = " in rep
assert 'Within(dayFinished, firstDay, lastDay)' in rep
for part in ["РОЗШИФРУВАННЯ ВИТРАТ ЗА КАТЕГОРІЯМИ","ЕФЕКТИВНІСТЬ АВТОПАРКУ","ДЕТАЛІЗАЦІЯ ЗАВЕРШЕНИХ РЕЙСІВ"]:
    assert part in rep
assert "Public Function TC_EquivalentUSD(" in fx
assert 'Array("PLN", "USD", "EUR", "UAH")' in fx
assert "DriverOverview live," in fx and "DriverOverview history," in fx
for part in ["Workbook_Activate()", "Workbook_Deactivate()", "Workbook_BeforeClose(", "Workbook_WindowDeactivate(", "RibbonForTransport", "RibbonRestore"]:
    assert part in ev
assert not any(x in ev for x in ("req.Open","TC_RatesFetchNBP","Invoke-WebRequest"))
assert "$workbookComponent.CodeModule.AddFromString($eventSource)" in inst
assert inst.index("$book.SaveAs($target,52)") < inst.index("$book.Close($false)")
assert "codeload.github.com" not in inst
for name,code in [("Insights",rep),("FX",fx),("Events",ev)]:
    for token in ("Sub","Function"):
        starts=len(re.findall(r"^\\s*(?:Public |Private |Friend )?"+token+r"\\s+\\w+\\s*\\(",code,re.M))
        ends=len(re.findall(r"^\\s*End "+token+r"\\s*$",code,re.M))
        assert starts == ends,(name,token,starts,ends)
print("V10.12 static: detailed reports, driver dashboards, ribbon restore, offline build PASS")
