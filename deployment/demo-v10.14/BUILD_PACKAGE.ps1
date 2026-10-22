#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$OutputZip = (Join-Path (Join-Path (Join-Path $PSScriptRoot '..\..') 'dist') 'Transport_Center_DEMO_v10.14.7_FAST_REFRESH_FULL.zip'),
    [string]$BaseWorkbook = (Join-Path $PSScriptRoot 'Транспортний_центр_v8_робочий_Times_New_Roman.xlsx')
)
$ErrorActionPreference = 'Stop'
$python = Get-Command python.exe -ErrorAction SilentlyContinue
if (-not $python) { throw 'Python 3 required for rebuilding the ZIP locally. The downloaded ready ZIP requires no Python.' }
if (-not (Test-Path -LiteralPath $BaseWorkbook -PathType Leaf)) {
    Write-Host 'Full ZIP requires an XLSX: creating isolated fictional baseline locally with Microsoft Excel...' -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot 'CREATE_DEMO_BASE.ps1') -OutputFile $BaseWorkbook
    if (-not (Test-Path -LiteralPath $BaseWorkbook -PathType Leaf)) { throw 'XLSX was not created' }
}
& $python.Source (Join-Path $PSScriptRoot 'ASSEMBLE_PACKAGE.py') --out $OutputZip --base-workbook $BaseWorkbook --require-base
if ($LASTEXITCODE -ne 0) { throw "Full ZIP assembly failed: $LASTEXITCODE" }
Write-Host "Complete offline demo ZIP with included XLSX: $OutputZip" -ForegroundColor Green
