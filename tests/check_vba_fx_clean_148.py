"""Guard against VBA Currency keyword reappearing as an identifier in CLEAN FX helpers."""
from pathlib import Path
import re

source = Path("deployment/demo-v10.14/TransportStatusFX.bas").read_text(encoding="utf-8-sig")
vba_files = [
    Path("deployment/demo-v10.14/TransportStatusFX.bas"),
    Path("deployment/demo-v10.14/TransportInsights.bas"),
    Path("deployment/offline-v10.9/TransportApp.bas"),
    Path("vba/TransportCenter.bas"),
]

reserved_identifiers = {"currency", "date", "integer", "long", "single", "double", "string", "boolean", "variant", "object", "sub", "function"}
for file in vba_files:
    code = file.read_text(encoding="utf-8-sig")
    for match in re.finditer(r"\b(?:ByVal|ByRef|Dim)\s+(\w+)\s+As\s+", code, re.I):
        token = match.group(1).casefold()
        assert token not in reserved_identifiers, (
            f"{file}:{code.count(chr(10), 0, match.start()) + 1}: "
            f"reserved VBA identifier '{match.group(1)}' in declaration"
        )

for required in (
    "Public Function TC_USDFrom(ByVal amount As Double, ByVal currencyCode As String) As Double",
    "Public Function TC_PLNPerUnit(ByVal currencyCode As String) As Double",
    'currencyCode = UCase$(Trim$(currencyCode))',
    'TC_PLNPerUnit = Rate(UCase$(Trim$(currencyCode)))',
):
    assert required in source, required

assert "ByVal currency As String" not in source
assert "ByRef currency As String" not in source

app = Path("deployment/offline-v10.9/TransportApp.bas").read_text(encoding="utf-8-sig")
assert "TransportStatusFX.TC_USDFrom(amount, cur)" in app
assert "TransportStatusFX.TC_PLNPerUnit(cur)" in app
print("PASS: Currency reserved-name syntax regression blocked; caller contracts unchanged")
print("NOTE: This is static analysis, not an Excel VBA compile.")
