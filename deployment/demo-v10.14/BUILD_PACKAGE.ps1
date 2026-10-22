#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$OutputZip = (Join-Path (Join-Path (Join-Path $PSScriptRoot '..\..') 'dist') 'Transport_Center_DEMO_v10.14.zip')
)
$ErrorActionPreference = 'Stop'
$python = Get-Command python.exe -ErrorAction SilentlyContinue
if(-not $python){ throw 'Python 3 is required only to reassemble the development ZIP. The end user needs only the distributed ZIP and Microsoft Excel.' }
$source = Join-Path $PSScriptRoot 'ASSEMBLE_PACKAGE.py'
& $python.Source $source --out $OutputZip
if($LASTEXITCODE -ne 0){ throw "Demo ZIP assembly failed: $LASTEXITCODE" }
Write-Host "Self-contained synthetic installer ZIP: $OutputZip" -ForegroundColor Green
