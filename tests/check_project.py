"""Static project contract tests + optional read-only XLSX validation.
No Office process is required. These tests DO NOT execute VBA.
"""
from pathlib import Path
from zipfile import ZipFile
import re
import sys
from xml.etree import ElementTree

ROOT = Path(__file__).resolve().parents[1]
VBA = (ROOT / "vba" / "TransportCenter.bas").read_text(encoding="utf-8-sig")
INSTALL = (ROOT / "INSTALL.ps1").read_text(encoding="utf-8-sig")


def test_vba():
    names = [
        "ЗаповнитиДані", "ЗберегтиДані", "ДодатиВодія", "ПерейтиДоЗвітів",
        "РезервнаКопія", "ЗакритиРейс", "ОновитиЗІнтернету"
    ]
    for name in names:
        assert len(re.findall(r"Public Sub\s+" + name + r"\s*\(\)", VBA)) == 1, name
    sections = re.split(r"(?=Public Sub\s+)", VBA)
    web = [part for part in sections if any(token in part for token in ("ServerXMLHTTP", ".Open \"GET\"", "api.nbp.pl/api/"))]
    assert len(web) == 1 and web[0].startswith("Public Sub ОновитиЗІнтернету()"), "Network access outside explicit button"
    for forbidden in ("Workbook_Open(", "Workbook_SheetCalculate(", "Worksheet_Calculate(", "Worksheet_Change("):
        assert forbidden not in VBA, "Implicit event-based operation: " + forbidden
    assert "SaveCopyAs" in VBA and "_Резервні_копії" in VBA
    assert 'sh.Range("B5").Value = fullName' in VBA
    assert 'sh.Range("E5").Value = vehicle' in VBA
    assert 'f.Range("B6").Value = Date' in VBA
    assert 'a.Cells(q, 12).Value = "ЗАКРИТО"' in VBA
    assert 'Call BackupNow("before_close")' in VBA


def test_install():
    required = [
        "$excel.Calculation = -4135", # disable recalculation before source Open
        "$book.Worksheets.Open(",      # not actually used; check separately
    ]
    assert required[0] in INSTALL
    assert "$book = $excel.Workbooks.Open(" in INSTALL
    assert INSTALL.index("$excel.Calculation = -4135") < INSTALL.index("$book = $excel.Workbooks.Open(")
    for token in (
        "$sheet.Range('B5').Value2 = ''", "$sheet.Range('B6').Value2 = 0",
        "$sheet.Range('B7').Value2 = ''", "$sheet.Range('B8').Value2 = ''",
        "$rates.Range('G13:H52').ClearContents()",
        "=Реєстр_водіїв!$A$13:$A$52",
        "$form.Range('B5').Validation.Add(3,1,1,'=TC_Drivers')",
        "$book.SaveAs($destinationPath, 52)",
    ):
        assert token in INSTALL, token
    assert "$excel.Calculation = -4105" in INSTALL
    assert "Зміст" in INSTALL and "Worksheets.Item(1)" in INSTALL
    assert "Транспортний_центр.xlsm" in INSTALL
    # Code must not introduce Workbook events which call web network automatically.
    assert not re.search(r"Workbook_Open[\s\S]{0,100}ОновитиЗІнтернету", INSTALL)


def validate_xlsx(filename):
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(filename) as z:
        assert z.testzip() is None, "Corrupt ZIP"
        for f in z.namelist():
            if f.endswith((".xml", ".rels")):
                ElementTree.fromstring(z.read(f))
        wb = ElementTree.fromstring(z.read("xl/workbook.xml"))
        names = [s.attrib["name"] for s in wb.findall(".//m:sheets/m:sheet", ns)]
        assert names[0] == "Зміст"
        for n in ("Нарахування", "Реєстр_водіїв", "ШАБЛОН_ВОДІЯ",
                  "Курси_NBP", "NBP_дані", "Звіти", "Історія_тарифів"):
            assert n in names, "Missing worksheet: " + n
        charts = [p for p in z.namelist() if re.search(r"(^|/)chart\d+\.xml$", p)]
        assert len(charts) == 4, "Expected 4 chart objects"


if __name__ == "__main__":
    test_vba()
    test_install()
    if len(sys.argv) > 1:
        validate_xlsx(sys.argv[1])
        print("OK: static VBA/installer contracts; XLSX archive, XML, sheets and charts.")
    else:
        print("OK: static VBA/installer contracts. Excel COM/VBA runtime NOT tested.")
