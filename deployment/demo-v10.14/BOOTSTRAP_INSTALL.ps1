#Requires -Version 5.1
# v10.14.4: All Unicode filenames are interpreted by PowerShell, never cmd.exe.
# The script MUST be UTF-8 with BOM for Windows PowerShell 5.1.
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$base = Join-Path $root 'Транспортний_центр_v8_робочий_Times_New_Roman.xlsx'
$target = Join-Path $root 'Транспортний_центр.xlsm'
$installer = Join-Path $root 'INSTALL_OFFLINE.ps1'
$bootstrapLog = Join-Path $root 'BOOTSTRAP_ERROR.txt'
try {
    Write-Host 'Transport Center v10.14.4 / Windows PowerShell bootstrap' -ForegroundColor Cyan
    if (-not (Test-Path -LiteralPath $base -PathType Leaf)) {
        throw 'BASE_XLSX_MISSING: full ZIP was not extracted correctly. Unpack ALL files into a new directory.'
    }
    if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) {
        throw 'INSTALLER_MISSING: INSTALL_OFFLINE.ps1 was not found next to START_INSTALL.cmd.'
    }
    if (Test-Path -LiteralPath $target) {
        throw 'XLSM_EXISTS: an earlier test workbook is already here. Use a NEW EMPTY folder for this release. Existing data will never be overwritten.'
    }
    Write-Host 'Base XLSX found. Starting the full offline Excel installer...' -ForegroundColor Green
    & $installer
    if (-not $?) { throw 'INSTALL_SCRIPT_FAILED: consult INSTALL_ERROR.txt / VBA_SMOKE_TEST_FAILED.txt' }
    if (-not (Test-Path -LiteralPath $target -PathType Leaf)) {
        throw 'NO_OUTPUT_XLSM: installer returned without creating the macro workbook.'
    }
    Write-Host 'INSTALLER COMPLETE. This remains a DEMO build; test all Excel macros locally.' -ForegroundColor Green
    exit 0
} catch {
    $msg = $_.Exception.Message
    @(
        'Transport Center v10.14.4 bootstrap failure',
        ('Timestamp: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
        ('Directory: ' + $root),
        ('Reason: ' + $msg),
        'Do not delete or overwrite an existing XLSM.',
        'For VBA runtime errors also inspect VBA_SMOKE_TEST_FAILED.txt and INSTALL_ERROR.txt.'
    ) | Out-File -LiteralPath $bootstrapLog -Encoding UTF8
    Write-Host ('ERROR: ' + $msg) -ForegroundColor Red
    Write-Host ('Diagnostic file: ' + $bootstrapLog) -ForegroundColor Yellow
    exit 1
}