#!/usr/bin/env python3
"""Pack complete fictional v10.14 source+one-click Windows Excel installer as ZIP."""
import argparse
import hashlib
import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT=Path(__file__).resolve().parents[2]
PREFIX="Транспортний_центр_DEMO_v10.14"
FILES={
 "deployment/demo-v10.12/INSTALL_OFFLINE.ps1":"INSTALL_OFFLINE.ps1",
 "deployment/demo-v10.10/DEMO_SEED.ps1":"DEMO_SEED.ps1",
 "deployment/demo-v10.14/CREATE_DEMO_BASE.ps1":"CREATE_DEMO_BASE.ps1",
 "vba/TransportCenter.bas":"vba/TransportCenter.bas",
 "deployment/offline-v10.9/TransportApp.bas":"vba/TransportApp.bas",
 "deployment/demo-v10.13/TransportInsights.bas":"vba/TransportInsights.bas",
 "deployment/demo-v10.14/TransportStatusFX.bas":"vba/TransportStatusFX.bas",
 "deployment/demo-v10.12/TransportWorkbookEvents.txt":"vba/TransportWorkbookEvents.txt",
}
INSTALL_CMD=r"""@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo TRANSPORT CENTER v10.14 - SYNTHETIC DEMO
echo All transactions are fictional. Microsoft Excel for Windows is required.
echo.
if not exist "Транспортний_центр_v8_робочий_Times_New_Roman.xlsx" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0CREATE_DEMO_BASE.ps1"
  if errorlevel 1 goto error
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0INSTALL_OFFLINE.ps1"
if errorlevel 1 goto error
echo.
echo DONE. Open "Транспортний_центр.xlsm" and enable local trusted macros.
echo Test: Current state, Archive, Finance report, closing/reopening Excel.
echo.
pause
exit /b 0
:error
echo.
echo INSTALL FAILED. See INSTALL_ERROR.txt if created. Preserve the source book.
pause
exit /b 1
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
    args=parser.parse_args()
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
            z.writestr(f"{PREFIX}/START_INSTALL.cmd",INSTALL_CMD.encode("ascii"))
            z.writestr(f"{PREFIX}/ІНСТРУКЦІЯ_ТЕСТУВАННЯ.txt",README.encode("utf-8-sig"))
        with zipfile.ZipFile(out) as z:
            assert z.testzip() is None
            names=set(z.namelist())
            assert f"{PREFIX}/START_INSTALL.cmd" in names
            assert f"{PREFIX}/data/demo_seed.json" in names
            assert f"{PREFIX}/vba/TransportStatusFX.bas" in names
            assert f"{PREFIX}/vba/TransportInsights.bas" in names
        digest=hashlib.sha256(out.read_bytes()).hexdigest()
        print(f"READY {out} bytes={out.stat().st_size} SHA256={digest}")
if __name__=="__main__":
    main()
