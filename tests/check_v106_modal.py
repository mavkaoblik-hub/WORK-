"""Static contract for v10.6 modal form: prepared lists, no worksheet acquisition during activation."""
from pathlib import Path
import re

src = Path("deployment/offline-v10.6/TransportApp.bas").read_text(encoding="utf-8-sig")
assert src.count("Private Sub PrepareFormLists(") == 1
assert src.count("Public Sub TC_FormInit(") == 1
show = src.split("Private Sub ShowTC(", 1)[1].split("Private Sub PrepareFormLists(", 1)[0]
init = src.split("Public Sub TC_FormInit(", 1)[1].split("Private Sub AddTripChoice(", 1)[0]
prepare = src.split("Private Sub PrepareFormLists(", 1)[1].split("Public Sub TC_FormInit(", 1)[0]
assert show.index("PrepareFormLists stage") < show.index("frmTCAdmin.Show vbModal")
assert not re.search(r"\b(?:WS|Worksheets|Cells|Range)\s*\(", init), "Do not dereference workbook during UserForm_Activate"
assert "mUiLoading = True" in init and "mUiLoading = False" in init
assert "If mUiLoading Then Exit Sub" in src
for mode in ("NEW", "DRIVER", "CAR", "ROUTE", "PARKING", "START", "FINISH", "COST", "REPORT"):
    assert '"' + mode + '"' in prepare, mode
for part in ("mUiDrivers", "mUiVehicles", "mUiRoutes", "mUiParkings", "mUiTrips"):
    assert "Set " + part + " = New Collection" in prepare, part
    assert "For Each entry In " + part in init, part
assert "If IsError(v) Then Exit Function" in src
print("v10.6: pre-modal, mode-specific lists and event guards verified")
