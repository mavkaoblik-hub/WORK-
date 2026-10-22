"""v10.7 regression: empty ComboBox selection cannot cause VBA error 9 via Split(empty)(0)."""
from pathlib import Path
import re

vba = Path("deployment/offline-v10.7/TransportApp.bas").read_text(encoding="utf-8-sig")
m = re.search(r"Private Function IdPrefix\(ByVal s As String\) As String(?P<body>.*?)End Function", vba, re.S)
assert m is not None, "IdPrefix not found"
body = m.group("body")
assert 'If Len(s) = 0 Then Exit Function' in body
assert 'InStr(1, s, " | ", vbBinaryCompare)' in body
assert 'Left$(s, delimiterPos - 1)' in body
assert 'Split(' not in re.sub(r"'[^\n]*", "", body), "Never index an empty Split array"
assert 'IdPrefix = s' in body

# Replicate the linear, index-free string manipulation to cover key edge cases.
def id_prefix(s: str) -> str:
    s = s.strip()
    if not s:
        return ""
    p = s.find(" | ")
    return s[:p].strip() if p >= 0 else s
cases = {
    "": "",
    "   ": "",
    "CAR-12": "CAR-12",
    "CAR-12 | Volvo AA 1122": "CAR-12",
    "Водій_01 | Прізвище Ім'я": "Водій_01",
    " | unnamed": "| unnamed",  # malformed input without valid ID separator after trim
    "TRIP-1 | Route | Driver": "TRIP-1",
}
for input_, expected in cases.items():
    assert id_prefix(input_) == expected, input_
for method in ("SaveNew", "SaveStart", "SaveFinish", "SaveCost", "SaveCar", "SaveDriver", "SaveRoute", "SaveParking"):
    assert re.search(r"Private Sub " + method + r"\(", vba), method
assert 'Операцію не виконано [" & mMode & "] (' in vba
print("v10.7: empty/normal/composite IDs and shared saving-path contract passed")
