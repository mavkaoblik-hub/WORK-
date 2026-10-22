#Requires -Version 5.1
[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][string]$SourceWorkbook,
 [Parameter(Mandatory=$true)][string]$DemoSeedJson,
 [string]$OutputFolder = (Join-Path $PSScriptRoot '_build_v10.14')
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
function Copy-Required([string]$src,[string]$dst) {
 if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { throw "Missing package dependency: $src" }
 $parent = Split-Path -Parent $dst
 if (-not (Test-Path -LiteralPath $parent)) { [void](New-Item -ItemType Directory -Path $parent -Force) }
 Copy-Item -LiteralPath $src -Destination $dst -Force
}
if (-not (Test-Path -LiteralPath $SourceWorkbook -PathType Leaf)) { throw "Source XLSX not found: $SourceWorkbook" }
if (-not (Test-Path -LiteralPath $DemoSeedJson -PathType Leaf)) { throw "Demo fixture JSON not found: $DemoSeedJson" }
$seed = Get-Content -LiteralPath $DemoSeedJson -Raw -Encoding UTF8 | ConvertFrom-Json
if ($seed.format -ne 'TRANSPORT_DEMO_SEED_V1') { throw 'Invalid demo seed format; never inject unverified production data' }
if (Test-Path -LiteralPath $OutputFolder) { throw "Destination already exists; refusing to overwrite: $OutputFolder" }
[void](New-Item -ItemType Directory -Path $OutputFolder)
try {
 $f = @{ 
  'deployment/demo-v10.12/INSTALL_OFFLINE.ps1' = 'INSTALL_OFFLINE.ps1'
  'deployment/demo-v10.10/DEMO_SEED.ps1' = 'DEMO_SEED.ps1'
  'vba/TransportCenter.bas' = 'vba/TransportCenter.bas'
  'deployment/offline-v10.9/TransportApp.bas' = 'vba/TransportApp.bas'
  'deployment/demo-v10.13/TransportInsights.bas' = 'vba/TransportInsights.bas'
  'deployment/demo-v10.14/TransportStatusFX.bas' = 'vba/TransportStatusFX.bas'
  'deployment/demo-v10.12/TransportWorkbookEvents.txt' = 'vba/TransportWorkbookEvents.txt'
  'data/routes_europe.json' = 'data/routes_europe.json'
 }
 foreach ($name in $f.Keys) { Copy-Required (Join-Path $root $name) (Join-Path $OutputFolder $f[$name]) }
 Copy-Required $SourceWorkbook (Join-Path $OutputFolder 'Транспортний_центр_v8_робочий_Times_New_Roman.xlsx')
 Copy-Required $DemoSeedJson (Join-Path $OutputFolder 'data/demo_seed.json')
 $installer = Join-Path $OutputFolder 'INSTALL_OFFLINE.ps1'
 $content = [IO.File]::ReadAllText($installer, [Text.Encoding]::UTF8)
 if (-not $content.Contains("TransportStatusFX.bas") -or -not $content.Contains('DEMO_SEED.ps1')) { throw 'Installer contract mismatch' }
 $launcher = '@echo off' + "`r`n" + 'cd /d "%~dp0"' + "`r`n" + 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0INSTALL_OFFLINE.ps1"' + "`r`n" + 'pause' + "`r`n"
 [IO.File]::WriteAllText((Join-Path $OutputFolder 'START_INSTALL.cmd'),$launcher,[Text.Encoding]::ASCII)
 Write-Host "v10.14 demo package ready at: $OutputFolder" -ForegroundColor Green
 Write-Host 'Run START_INSTALL.cmd only on Windows with licensed Microsoft Excel. No GitHub access required during installation.'
} catch {
 Write-Warning "Incomplete package left for inspection: $OutputFolder"
 throw
}
