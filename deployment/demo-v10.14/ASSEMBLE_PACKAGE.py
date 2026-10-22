#!/usr/bin/env python3
"""Package synthetic demo source; when given a verified XLSX, also deliver a FULL installer ZIP."""
import argparse
import hashlib
import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT=Path(__file__).resolve().parents[2]
PREFIX="Транспортний_центр_DEMO_v10.14.7"
FILES={
 "deployment/demo-v10.14/INSTALL_DIAGNOSTIC.ps1":"INSTALL_OFFLINE.ps1",
 "deployment/demo-v10.10/DEMO_SEED.ps1":"DEMO_SEED.ps1",
 "deployment/demo-v10.14/CREATE_DEMO_BASE.ps1":"CREATE_DEMO_BASE.ps1",
 "deployment/demo-v10.14/BOOTSTRAP_INSTALL.ps1":"BOOTSTRAP_INSTALL.ps1",
 "vba/TransportCenter.bas":"vba/TransportCenter.bas",
 "deployment/offline-v10.9/TransportApp.bas":"vba/TransportApp.bas",
 "deployment/demo-v10.13/TransportInsights.bas":"vba/TransportInsights.bas",
 "deployment/demo-v10.14/TransportStatusFX.bas":"vba/TransportStatusFX.bas",
 "deployment/demo-v10.14/TransportWorkbookEvents.txt":"vba/TransportWorkbookEvents.txt",
}
INSTALL_CMD=r"""@echo off
setlocal
REM ASCII-only launcher; Unicode filenames belong in BOOTSTRAP_INSTALL.ps1.
pushd "%~dp0"
if errorlevel 1 goto location_error
echo.
echo Transport Center DEMO v10.14.7 - Windows Excel offline installer
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\\BOOTSTRAP_INSTALL.ps1"
set "INSTALL_EXIT=%ERRORLEVEL%"
popd
if not "%INSTALL_EXIT%"=="0" goto installation_error
echo.
echo Installation script finished. Check the generated workbook in this folder.
pause
exit /b 0
:installation_error
echo.
echo INSTALLATION FAILED - see BOOTSTRAP_ERROR.txt, INSTALL_ERROR.txt or VBA_SMOKE_TEST_FAILED.txt.
echo Existing XLSM workbooks are NEVER overwritten.
pause
exit /b %INSTALL_EXIT%
:location_error
echo ERROR: Cannot open the extracted installation directory.
pause
exit /b 2
"""
README="""ТРАНСПОРТНИЙ ЦЕНТР v10.14 — ТЕСТОВА ОФЛАЙН-ЗБІРКА
22.10.2026 | УСІ ДАНІ ПОВНІСТЮ ВИГАДАНІ

ШВИДКИЙ СТАРТ (WINDOWS + EXCEL LTSC / DESKTOP):
1. Розпакуйте весь архів до окремої папки, наприклад C:\\TC_DEMO_v10_14.
2. Повністю закрийте Excel. Двічі натисніть START_INSTALL.cmd.
3. Для інсталяції потрібен одноразовий дозвіл Excel на доступ до моделі VBA:
   Файл -> Параметри -> Центр безпеки -> Параметри центру безпеки ->
   Параметри макросів -> Довіряти доступу до об'єктної моделі проєктів VBA.
4. Дочекайтеся появи Транспортний_центр.xlsm у цій папці.
5. Вимкніть зазначений дозвіл після успішного встановлення.
6. Відкрийте Транспортний_центр.xlsm, дозволяйте макроси тільки для
   цього перевіреного локального тестового файлу.
7. Перевірте: Поточний стан / Архів рейсів / Фінансовий звіт /
   автоматичне приховування і відновлення стрічки Excel.
8. Після перевірки закрийте книгу. Переконайтеся, що інша книга Excel
   відображає свою стрічку нормально.

Комплект уже містить 7 вигаданих водіїв, 9 вигаданих авто,
75 тестових рейсів (70 завершено, 3 активні, 2 заплановані)
та 259 вигаданих записів витрат.
Початкова XLSX створюється автоматично локально при першій установці.
Справжні дані та вихідна книга v8 НЕ потрібні.
Під час інсталяції НЕ здійснюються звернення до GitHub.
Єдина окрема інтернет-функція у книзі — ручна кнопка курсу NBP.

ЦЕ ТЕСТОВА ЗБІРКА. Не використовуйте її як основний облік.
Не копіюйте реальні записи у публічний репозиторій.

Якщо інсталятор зупинився, надішліть лише текст INSTALL_ERROR.txt;
не надсилайте персональні дані або реальні облікові книги без потреби.
GitHub перевіряє код та цілісність пакета; VBA на Excel LTSC
потрібно окремо перевірити на вашому ПК.
"""
def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("--out",default=str(ROOT/"dist"/(PREFIX+".zip")))
    parser.add_argument("--overwrite",action="store_true")
    parser.add_argument("--base-workbook",help="Prebuilt synthetic XLSX to include in the FULL installer ZIP")
    parser.add_argument("--require-base",action="store_true",help="Fail if no XLSX is supplied; never mislabel incomplete releases")
    args=parser.parse_args()
    base=Path(args.base_workbook).resolve() if args.base_workbook else None
    if args.require_base and base is None:
        raise SystemExit("Full release requires --base-workbook: missing XLSX, refusing incomplete archive")
    if base is not None:
        if not base.is_file() or base.suffix.lower()!=".xlsx":
            raise SystemExit(f"Missing valid XLSX path: {base}")
        import xml.etree.ElementTree as ET
        with zipfile.ZipFile(base) as wb:
            if wb.testzip() is not None or "xl/workbook.xml" not in wb.namelist():
                raise SystemExit("Corrupt/invalid base XLSX")
            root=ET.fromstring(wb.read("xl/workbook.xml"))
            ns={"s":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            names={tag.attrib["name"] for tag in root.findall(".//s:sheet",ns)}
            required={"Зміст","Нарахування","NBP_дані","Курси_NBP","Налаштування","Інструкція","Паркінги","Реєстр_водіїв","ШАБЛОН_ВОДІЯ","Водій_01","Водій_02","КІНЕЦЬ_ВОДІЇВ","Звіти"}
            if not required.issubset(names):
                raise SystemExit(f"Base XLSX missing expected sheets: {sorted(required-names)}")
    out=Path(args.out).resolve()
    out.parent.mkdir(parents=True,exist_ok=True)
    if out.exists() and not args.overwrite:
        raise SystemExit(f"Refusing existing output: {out}")
    with tempfile.TemporaryDirectory() as tmp:
        seed=Path(tmp)/"demo_seed.json"
        subprocess.run([sys.executable, str(ROOT/"deployment/demo-v10.14/GENERATE_SEED.py"),
                        "--routes",str(ROOT/"data/routes_europe.json"),"--out",str(seed)],check=True)
        with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED,compresslevel=9) as z:
            for src,dst in FILES.items():
                path=ROOT/src
                if not path.is_file():
                    raise FileNotFoundError(path)
                z.write(path,arcname=f"{PREFIX}/{dst}")
            raw=json.loads((ROOT/"data/routes_europe.json").read_text(encoding="utf-8-sig"))
            listed=raw["routes"] if isinstance(raw,dict) else raw
            if len(listed)!=114: raise ValueError("Route catalogue is not 114 entries")
            normalized=[{
                "id":v["id"],"start":v.get("start",v.get("from")),
                "end":v.get("end",v.get("to")),"km":v["km"],
                "hours":v.get("hours"),"typ":v.get("typ",v.get("kind","")),
                "source":v.get("source",""),"year":v.get("year",v.get("source_year",2026)),
                "note":v.get("note",v.get("remark",""))
            } for v in listed]
            z.writestr(f"{PREFIX}/data/routes_europe.json",
                       json.dumps(normalized,ensure_ascii=False,indent=2).encode("utf-8"))
            z.write(seed,arcname=f"{PREFIX}/data/demo_seed.json")
            if base is not None:
                z.write(base,arcname=f"{PREFIX}/Транспортний_центр_v8_робочий_Times_New_Roman.xlsx")
            z.writestr(f"{PREFIX}/START_INSTALL.cmd",INSTALL_CMD.encode("utf-8"))
            z.writestr(f"{PREFIX}/ІНСТРУКЦІЯ_ТЕСТУВАННЯ.txt",README.encode("utf-8-sig"))
        with zipfile.ZipFile(out) as z:
            assert z.testzip() is None
            names=set(z.namelist())
            assert f"{PREFIX}/START_INSTALL.cmd" in names
            assert f"{PREFIX}/BOOTSTRAP_INSTALL.ps1" in names
            launcher=z.read(f"{PREFIX}/START_INSTALL.cmd")
            if not launcher.isascii():
                raise RuntimeError("CMD must remain ASCII-only, all Unicode paths must be handled in PowerShell")
            if b"BOOTSTRAP_INSTALL.ps1" not in launcher:
                raise RuntimeError("CMD is not delegated to Unicode-safe PowerShell bootstrap")
            assert f"{PREFIX}/data/demo_seed.json" in names
            assert f"{PREFIX}/vba/TransportStatusFX.bas" in names
            assert f"{PREFIX}/vba/TransportInsights.bas" in names
            assert f"{PREFIX}/vba/TransportWorkbookEvents.txt" in names
            fx_module=z.read(f"{PREFIX}/vba/TransportStatusFX.bas").decode("utf-8-sig")
            events=z.read(f"{PREFIX}/vba/TransportWorkbookEvents.txt").decode("utf-8-sig")
            assert "detailRegion.Value2 = detailValues" in fx_module
            assert "monthlyHeader = Application.Max(20" in fx_module
            assert "Private mStatusLastRefresh As Date" in events
            if base is not None:
                assert f"{PREFIX}/Транспортний_центр_v8_робочий_Times_New_Roman.xlsx" in names
        digest=hashlib.sha256(out.read_bytes()).hexdigest()
        print(f"READY {out} bytes={out.stat().st_size} SHA256={digest}")
        print("FULL INSTALLER + BASE XLSX INCLUDED" if base else
              "SOURCE-ONLY ZIP: NO BASE XLSX! Use --base-workbook --require-base for actual full delivery.")
if __name__=="__main__":
    main()
