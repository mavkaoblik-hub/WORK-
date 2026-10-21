param([string]$Workbook = "Транспортний_центр.xlsm")
$ErrorActionPreference='Stop'
$dir=Split-Path -Parent $MyInvocation.MyCommand.Path
$target=Join-Path $dir $Workbook
$folder=Join-Path $dir "_Резервні_копії"
if (!(Test-Path -LiteralPath $folder)) { throw "Немає папки резервних копій: $folder" }
$items=@(Get-ChildItem -LiteralPath $folder -Filter '*.xlsm' | Sort-Object LastWriteTime -Descending)
if ($items.Count -eq 0) { throw "Немає резервних копій." }
Write-Host 'Виберіть копію для відновлення (Excel має бути закритий):'
for($i=0;$i -lt $items.Count;$i++) { Write-Host ("{0}. {1} ({2})" -f ($i+1), $items[$i].Name, $items[$i].LastWriteTime) }
$choice=Read-Host 'Номер копії'
$index=0
if (!( [int]::TryParse($choice,[ref]$index)) -or $index -lt 1 -or $index -gt $items.Count) { throw 'Невірний номер.' }
$confirm=Read-Host ("ВІДНОВИТИ " + $items[$index-1].Name + "? Введіть ТАК")
if ($confirm -cne 'ТАК') { Write-Host 'Скасовано'; exit }
if(Test-Path -LiteralPath $target) {
  # Exclusive handle check: do not overwrite a book currently opened by Excel.
  $handle = $null
  try { $handle = [System.IO.File]::Open($target, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None) }
  catch { throw "Закрийте Excel і повторіть відновлення. Робочий файл зараз зайнятий." }
  finally { if ($null -ne $handle) { $handle.Dispose() } }
  $before=Join-Path $folder ("before_restore_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".xlsm")
  Copy-Item -LiteralPath $target -Destination $before -ErrorAction Stop
  Write-Host "Поточну версію збережено: $before"
}
Copy-Item -LiteralPath $items[$index-1].FullName -Destination $target -Force -ErrorAction Stop
Write-Host "Відновлено: $target"
