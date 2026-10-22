# Demo v10.14 — live operations and archive analytics

**Development branch:** `feature/v10.14-operational-analytics`.

This version keeps the v10.13 detailed PDF engine while extending the v10.12 current-state/archive VBA module. **It is not a compiled or Excel-verified XLSM.** GitHub Actions performs source-level checks only.

## What changes
- **Поточний стан:** financial snapshot of trips with status `У рейсі`; explicitly separates *planned* USD revenue from *recorded* USD expense. It is not booked profit.
- **Поточний стан:** attention signals for missing odometer, at least 14 days underway, or recorded expenses exceeding planned revenue.
- **Архів рейсів:** completion-month revenue, all linked trip expenses, received payments, outstanding amount, actual mileage, margin and totals, plus an optional chart.
- Archive monthly expense attribution follows the **completion month** of the trip. PDF v10.13 expense details are instead assigned by **receipt/transaction date**. Thus monthly results across the two views intentionally differ.
- Chart deletion is performed on refresh to prevent old charts stacking up.

## Building on Windows
1. Check out this branch to a local directory.
2. Supply the original `Транспортний_центр_v8_робочий_Times_New_Roman.xlsx` outside the public repository.
3. Supply a **synthetic-only** JSON fixture with `format = TRANSPORT_DEMO_SEED_V1`; do not upload customer transactions or personal driver data.
4. Run:
   ```powershell
   .\deployment\demo-v10.14\BUILD_PACKAGE.ps1 -SourceWorkbook "C:\local\source.xlsx" -DemoSeedJson "C:\local\demo_seed.json"
   ```
5. In the resulting `_build_v10.14` directory run `START_INSTALL.cmd` in Windows with a licensed desktop Microsoft Excel and VBA project access temporarily enabled for installation.
6. Test opening/closing with another Excel workbook already open (ribbon should be restored), financial totals, blank datasets, monthly reconciliation, charts and PDF. Back up any real data separately.

**Important:** The installer inherits the existing demo-v10.12 offline build process. Its `DEMO_SEED.ps1` expects the full compatible seeded dataset. The packaging script validates the format and required paths but does not claim that arbitrary seeds will pass the older fixture validation. **Do not overwrite an existing XLSM**; the installer intentionally refuses.

Run `python tests/check_v114_demo.py` to perform static contract checks. No PowerShell or Excel execution has been performed here.
