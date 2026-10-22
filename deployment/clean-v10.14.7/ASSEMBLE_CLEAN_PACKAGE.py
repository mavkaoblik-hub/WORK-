#!/usr/bin/env python3
"""Reproducibly assemble a CLEAN v10.14.7 FULL installer with a verified EMPTY XLSX."""
import argparse
import hashlib
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PREFIX = "Транспортний_центр_CLEAN_v10.14.7"
NEEDED_SHEETS = {
    "Зміст", "Нарахування", "NBP_дані", "Курси_NBP",
    "Налаштування", "Інструкція", "Паркінги", "Реєстр_водіїв",
    "ШАБЛОН_ВОДІЯ", "КІНЕЦЬ_ВОДІЇВ", "Звіти"
}
FILES = {
    "deployment/clean-v10.14.7/INSTALL_CLEAN.ps1": "INSTALL_OFFLINE.ps1",
    "deployment/clean-v10.14.7/CREATE_CLEAN_BASE.ps1": "CREATE_CLEAN_BASE.ps1",
    "deployment/demo-v10.14/BOOTSTRAP_INSTALL.ps1": "BOOTSTRAP_INSTALL.ps1",
    "vba/TransportCenter.bas": "vba/TransportCenter.bas",
    "deployment/offline-v10.9/TransportApp.bas": "vba/TransportApp.bas",
    "deployment/demo-v10.13/TransportInsights.bas": "vba/TransportInsights.bas",
    "deployment/demo-v10.14/TransportStatusFX.bas": "vba/TransportStatusFX.bas",
    "deployment/demo-v10.14/TransportWorkbookEvents.txt": "vba/TransportWorkbookEvents.txt",
}
CMD = r"""@echo off
setlocal
pushd "%~dp0"
if errorlevel 1 exit /b 2
echo Transport Center CLEAN v10.14.7 - offline Windows Excel installation
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\BOOTSTRAP_INSTALL.ps1"
set "RC=%ERRORLEVEL%"
popd
if not "%RC%"=="0" goto bad
echo Installer finished. Open Transport Center XLSM from the extracted directory.
pause
exit /b 0
:bad
echo Installation failed - see BOOTSTRAP_ERROR.txt / INSTALL_ERROR.txt.
pause
exit /b %RC%
"""
README = """ТРАНСПОРТНИЙ ЦЕНТР CLEAN v10.14.7
Чиста збірка без водіїв, автомобілів, рейсів, витрат, паркінгів та операцій.
Довідник 114 маршрутів збережено (це не облікові транзакції).
Розпакувати ZIP у НОВУ папку Windows, закрити Excel,
запустити START_INSTALL.cmd. Інсталятор сформує XLSM, не перезаписуючи існуючі файли.
Може знадобитися тимчасове дозволення доступу до моделі VBA в Excel Trust Center.
Після встановлення вимкніть цей дозвіл.
Залишено оптимізації v10.14.7, PDF, фінансову звітність, резервні копії, усі форми.
Excel LTSC та VBA виконуються і перевіряються локально на Windows.
"""

def check_workbook(p):
    if not p.is_file():
        raise ValueError("Prebuilt empty XLSX missing. Run CREATE_CLEAN_BASE.ps1 on Windows.")
    with zipfile.ZipFile(p) as z:
        assert z.testzip() is None
        book = ET.fromstring(z.read("xl/workbook.xml"))
        ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        names={e.attrib["name"] for e in book.findall(".//s:sheet", ns)}
        if names != NEEDED_SHEETS:
            raise ValueError("Clean book must have exactly 11 required sheets and no phantom driver tabs")
        xmltext = " ".join(z.read(n).decode("utf-8", errors="replace")
                           for n in z.namelist() if n.endswith(".xml"))
        for forbidden in ("DEMO-PARK-", "ДЕМОНСТРАЦІЙНИЙ ВОДІЙ", "Водій_01", "Водій_02"):
            if forbidden in xmltext:
                raise ValueError("Sample data found in allegedly empty base XLSX: " + forbidden)

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("--base-workbook", required=True)
    parser.add_argument("--out", default=str(ROOT / "dist" / "Transport_Center_CLEAN_v10.14.7_FULL.zip"))
    args=parser.parse_args()
    base=Path(args.base_workbook).resolve()
    check_workbook(base)
    out=Path(args.out).resolve()
    if out.exists():
        raise ValueError("Refusing to overwrite existing ZIP: " + str(out))
    out.parent.mkdir(parents=True, exist_ok=True)
    content={}
    for src,dst in FILES.items():
        b=(ROOT/src).read_bytes()
        content[dst]=b
    installer=content["INSTALL_OFFLINE.ps1"].decode("utf-8-sig")
    assert "DEMO_SEED.ps1" not in installer and "demo_seed.json" not in installer
    assert "CLEAN DATA CHECK" in installer
    startup=content["BOOTSTRAP_INSTALL.ps1"].decode("utf-8-sig")
    startup=startup.replace("This remains a DEMO build; test all Excel macros locally.",
                            "Empty book installed; verify local Excel functions.")
    content["BOOTSTRAP_INSTALL.ps1"]=b"\xef\xbb\xbf"+startup.encode("utf-8")
    assert CMD.isascii()
    content["START_INSTALL.cmd"]=CMD.encode("ascii")
    content["Транспортний_центр_v8_робочий_Times_New_Roman.xlsx"]=base.read_bytes()
    routes=json.loads((ROOT/"data/routes_europe.json").read_text(encoding="utf-8-sig"))
    routes=routes["routes"] if isinstance(routes,dict) else routes
    assert len(routes)==114
    normalized=[{
        "id":v["id"],"start":v.get("start",v.get("from")),
        "end":v.get("end",v.get("to")),"km":v["km"],
        "hours":v.get("hours"),"typ":v.get("typ",v.get("kind","")),
        "source":v.get("source",""),"year":v.get("year",v.get("source_year",2026)),
        "note":v.get("note",v.get("remark",""))
    } for v in routes]
    content["data/routes_europe.json"]=json.dumps(normalized,ensure_ascii=False,indent=2).encode("utf-8")
    content["ІНСТРУКЦІЯ.txt"]=README.encode("utf-8-sig")
    manifest="SHA256 / payload files\n"
    for n,b in sorted(content.items()):
        manifest+=hashlib.sha256(b).hexdigest()+"  "+n+"\n"
    content["SHA256_MANIFEST.txt"]=manifest.encode("utf-8")
    with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED,compresslevel=9) as z:
        for name,blob in content.items():
            z.writestr(PREFIX+"/"+name,blob)
    with zipfile.ZipFile(out) as z:
        assert z.testzip() is None
        assert len(z.namelist())==len(content)
        assert not any("demo_seed" in n.lower() or "DEMO_SEED" in n for n in z.namelist())
    print("PASS: CLEAN FULL ZIP", out, out.stat().st_size)

if __name__=="__main__":
    main()
