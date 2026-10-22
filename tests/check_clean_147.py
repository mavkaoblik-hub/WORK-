"""CLEAN release contract: no fixtures/phantom data, keep the core application."""
from pathlib import Path
import re

d=Path("deployment/clean-v10.14.7")
installer=(d/"INSTALL_CLEAN.ps1").read_text(encoding="utf-8-sig")
base=(d/"CREATE_CLEAN_BASE.ps1").read_text(encoding="utf-8-sig")
build=(d/"ASSEMBLE_CLEAN_PACKAGE.py").read_text(encoding="utf-8-sig")
for t in ("DEMO_SEED.ps1","demo_seed.json","$cnt.Range('A6').Value2='Водій_01'","$cnt.Range('A7').Value2='Водій_02'"):
    assert t not in installer, f"Demo injection remains: {t}"
for k in ("TC_Рейси","TC_Витрати","Автопарк","Контакти_водіїв","Реєстр_водіїв","TC_Паркінги","Журнал_операцій"):
    assert k in installer, f"missing clean check: {k}"
assert "CLEAN DATA CHECK" in installer
assert installer.index("$book.SaveAs($target,52)") < installer.index("[void]$excel.Run($runName)")
assert "Транспортний_центр.xlsm вже існує" in installer
for fake in ("DEMO-PARK", "ДЕМОНСТРАЦІЙНИЙ ВОДІЙ", "'Водій_01'", "'Водій_02'", "Range('A13').Value2='Водій_01'"):
    assert fake not in base, fake
assert "'ШАБЛОН_ВОДІЯ'" in base and "'КІНЕЦЬ_ВОДІЇВ'" in base
assert "SaveAs($OutputFile,51)" in base
assert "NEEDED_SHEETS" in build and "no phantom driver" in build.lower()
assert "DEMO_SEED.ps1" in build, "explicit demo-fixture archive exclusion check required"
assert 'assert CMD.isascii()' in build
assert "source.xlsx" not in build.lower()
print("PASS: clean v10.14.7: empty workbook, no dummy drivers/parks, no demo seeding, versioned installer")
