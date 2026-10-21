#Requires -Version 5.1
param()
$ErrorActionPreference = 'Stop'
$homeDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceBook = Join-Path $homeDir 'Транспортний_центр_v8_робочий_Times_New_Roman.xlsx'
$outputBook = Join-Path $homeDir 'Транспортний_центр.xlsm'
# Pinned successful CI revision. Nothing in this launcher updates business data over the Internet.
$revision = '222350b20b2a3dbe38b05805bdaba4b6e6bb03bc'
$download = 'https://codeload.github.com/mavkaoblik-hub/WORK-/zip/' + $revision
$zipFile = Join-Path $homeDir '_transport_sources.zip'
$sourceRoot = Join-Path $homeDir '_transport_sources'
function Ensure-True($value,[string]$message) { if(-not $value) { throw $message } }
try {
    Ensure-True (Test-Path -LiteralPath $sourceBook) 'Не знайдено вхідної Excel-книги поруч із інсталятором.'
    Ensure-True (-not (Test-Path -LiteralPath $outputBook)) 'Транспортний_центр.xlsm вже існує! Оригінал не перезаписується. Збережіть його в іншому місці перед повторним встановленням.'
    Write-Host 'Транспортний центр: автоматична збірка VBA' -ForegroundColor Cyan
    Write-Host '1/4: завантаження перевіреного коду макросів із GitHub...'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $download -OutFile $zipFile -UseBasicParsing -TimeoutSec 90
    Ensure-True ((Get-Item -LiteralPath $zipFile).Length -gt 10000) 'Пакет із GitHub виглядає неповним.'
    if(Test-Path -LiteralPath $sourceRoot) { Remove-Item -LiteralPath $sourceRoot -Force -Recurse }
    Expand-Archive -LiteralPath $zipFile -DestinationPath $sourceRoot -Force
    $project = @(Get-ChildItem -LiteralPath $sourceRoot -Directory | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'INSTALL.ps1') })
    Ensure-True ($project.Count -eq 1) 'Структура пакета коду пошкоджена або змінилася.'
    $projectDir = $project[0].FullName
    foreach($path in @('INSTALL.ps1','CATALOGS.ps1','vba\TransportCenter.bas','vba\TransportCatalogs.bas','data\routes_europe.json','RESTORE_BACKUP.ps1','RESTORE_BACKUP.cmd')) {
        Ensure-True (Test-Path -LiteralPath (Join-Path $projectDir $path)) ('Відсутній файл: '+$path)
    }
    Write-Host '2/4: вбудовування VBA та створення кнопок через настільний Microsoft Excel...'
    Copy-Item -LiteralPath $sourceBook -Destination (Join-Path $projectDir (Split-Path -Leaf $sourceBook)) -ErrorAction Stop
    # Run INSTALL.ps1 INSIDE the current PowerShell session.
    # Starting a second powershell.exe can be blocked by Windows policy (Access is denied).
    $installScript = Join-Path $projectDir 'INSTALL.ps1'
    & $installScript -Source (Split-Path -Leaf $sourceBook) -Destination 'Транспортний_центр.xlsm'
    # Native LASTEXITCODE is not applicable to a PowerShell script. Exceptions propagate to catch.
    $builtBook = Join-Path $projectDir 'Транспортний_центр.xlsm'
    Ensure-True (Test-Path -LiteralPath $builtBook) 'Excel не створив Транспортний_центр.xlsm.'
    Write-Host '3/4: перевірка структури готового файлу...'
    Add-Type -AssemblyName System.IO.Compression
    $stream = [System.IO.File]::OpenRead($builtBook)
    try {
       $archive = [System.IO.Compression.ZipArchive]::new($stream,[System.IO.Compression.ZipArchiveMode]::Read)
       try {
           $names = @($archive.Entries | ForEach-Object { $_.FullName })
           Ensure-True ($names.Contains('xl/vbaProject.bin')) 'В готовому XLSM немає вбудованого VBA-проєкту.'
           Ensure-True ($names.Contains('xl/workbook.xml')) 'В готовому XLSM немає структури книги.'
           Ensure-True ($names.Contains('xl/worksheets/sheet1.xml')) 'В готовому XLSM немає першого аркуша.'
       } finally { $archive.Dispose() }
    } finally { $stream.Dispose() }
    Write-Host '4/4: переміщення готової макросної книги в основну папку...'
    Copy-Item -LiteralPath $builtBook -Destination $outputBook -ErrorAction Stop
    foreach($file in @('RESTORE_BACKUP.ps1','RESTORE_BACKUP.cmd')) {
       Copy-Item -LiteralPath (Join-Path $projectDir $file) -Destination (Join-Path $homeDir $file) -ErrorAction Stop
    }
    $backupDir = Join-Path $homeDir '_Резервні_копії'
    if(-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }
    Write-Host ('ГОТОВО! Макроси вбудовано автоматично: '+$outputBook) -ForegroundColor Green
    Write-Host 'Оригінальний XLSX збережений без змін.'
} catch {
    Write-Host ('ЗБІРКУ НЕ ЗАВЕРШЕНО: '+$_.Exception.Message) -ForegroundColor Red
    Write-Host 'Якщо помилка прямо згадує VBProject або програмний доступ до VBA, перевірте Excel > Файл > Параметри > Центр безпеки > Параметри центру безпеки > Параметри макросів > Довіряти доступу до програмної моделі проєктів VBA.' -ForegroundColor Yellow
    Write-Host 'Після збірки вимкніть цей тимчасовий дозвіл. Якщо помилка інша, надішліть повний текст.' -ForegroundColor Yellow
    exit 1
}
