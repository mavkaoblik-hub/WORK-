# Transport Center DEMO v10.14 — offline installer

**Status:** source-validated test release. The actual execution of VBA on customer Excel LTSC remains to be confirmed. The ZIP is generated in GitHub Actions and can be downloaded as an artifact.

## ZIP content and deployment

The entire `Transport_Center_DEMO_v10.14.zip` is **self-contained for a synthetic test**, including all VBA modules, the v10.12 installation engine, v10.13 detailed report module, v10.14 operational screen module, 114-route catalogue in installer format, 75 fictional trips and 259 fictional expense records.

1. Download the artifact from the **Build installable demo ZIP v10.14** workflow run.
2. Extract its inner `Transport_Center_DEMO_v10.14.zip` and unpack that ZIP into a **new, separate folder**.
3. On Windows with desktop Microsoft Excel, close all Excel windows and run `START_INSTALL.cmd`.
4. First launch creates a minimal, entirely synthetic `Транспортний_центр_v8_робочий_Times_New_Roman.xlsx` with Excel COM. It then invokes the unchanged compatible demo installer.
5. Temporarily allow Excel Trust access to VBA project object model for this installation if requested; disable it afterwards.
6. Open the created `Транспортний_центр.xlsm`; verify ribbons restore when switching to another workbook and on close, monthly archive, current state, finance PDF, and expense details.

**Windows PowerShell 5.1:** scripts containing Ukrainian worksheet names must be UTF-8 **with BOM**. The portable `CREATE_DEMO_BASE.ps1` is stored accordingly.

The ZIP is strictly a demonstration — it must not be used with production data or as a replacement for a user's previous accounting workbook. No unattended GitHub downloads occur during installation or workbook use. Only the explicit NBP refresh button may request FX data.

## Analytics

- **Поточний стан:** comparison of *planned* earnings and recorded cost for in-progress trips; missing odometer, 14+ days and overspend indicators.
- **Архів рейсів:** completion-month revenue, all linked expenses, payments, outstanding balance, mileage, margins, and trend chart.
- **PDF:** v10.13 receipts ledger and monthly breakdown continue unchanged.

**Important:** archive expense allocation uses the trip's completion month, while PDF expense reporting uses each receipt's date. Those are two intentionally different metrics.

## Developer build

Use Python 3 (standard library only):

```shell
python3 deployment/demo-v10.14/ASSEMBLE_PACKAGE.py --out dist/Transport_Center_DEMO_v10.14.zip
python3 tests/check_v114_demo.py
```

Or run `BUILD_PACKAGE.ps1` (Python 3 required only when rebuilding locally). An Excel license is not required to assemble the ZIP. It *is* required to install and execute it.
