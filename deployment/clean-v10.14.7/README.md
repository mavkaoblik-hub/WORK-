# CLEAN v10.14.7 — empty accounting system

A **clean initial installation**, separate from the demo fixtures. No transactions are transferred from previous workbooks.

## Delivery and included components

- Empty 11-sheet pre-install XLSX, including `Зміст`, `Нарахування`, `Реєстр_водіїв`, `ШАБЛОН_ВОДІЯ` and service tabs. No pre-created `Водій_01/02`.
- `INSTALL_CLEAN.ps1` builds the same application and macros as v10.14.7, but **does not call** `DEMO_SEED.ps1` and **does not add placeholder drivers**.
- Installer checks for ZERO records in `TC_Рейси`, `TC_Витрати`, `Автопарк`, `Контакти_водіїв`, `Реєстр_водіїв`, `TC_Паркінги`, `Журнал_операцій` before creating the XLSM.
- 114 standard route entries remain in the reference catalogue. These are **not** trips.
- Status/archive screens, optimized v10.14.7 VBA, PDF finance reports, forms and backup operations remain intact.
- No demo JSON, dummy vehicle, fake parking tariff, dummy exchange rate, or seeded income/expenses.

## Reproduce the user-facing FULL ZIP

On Windows with Microsoft Excel installed, run `CREATE_CLEAN_BASE.ps1` once to generate the initial empty XLSX, then:

```powershell
python deployment/clean-v10.14.7/ASSEMBLE_CLEAN_PACKAGE.py --base-workbook "deployment/clean-v10.14.7/Транспортний_центр_v8_робочий_Times_New_Roman.xlsx" --out "dist/Transport_Center_CLEAN_v10.14.7_FULL.zip"
```

The delivered ZIP in chat already includes the ready empty XLSX created with a spreadsheet engine, so **the user needs no Python**. Extract the ZIP to a fresh Windows folder, close Excel, run `START_INSTALL.cmd`. Trust access to VBA project may need temporary permission for installation; restore the safer Trust Center setting afterwards.

**Runtime caveat:** structural ZIP/OOXML checks are not a substitute for executing the installed macros in the user's Excel LTSC. The output XLSM is generated locally and is not overwritten automatically.

## Release checks

```bash
python tests/check_clean_147.py
```

The clean installer must never gain a dependency on demo fixture files; a future regression should fail CI.
