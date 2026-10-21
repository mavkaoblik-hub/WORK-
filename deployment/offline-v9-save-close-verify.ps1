# Offline v9: save/close Office before validating XLSM.
# Same ordering used in customer release Transport_Center_Offline_v9.zip.
# Usage: Complete-TCXlsmSave -Workbook $book -Excel $excel -FileName $target
# The caller must open Excel itself; this helper does NOT download anything.
function Complete-TCXlsmSave {
  param(
    [Parameter(Mandatory=$true)]$Workbook,
    [Parameter(Mandatory=$true)]$Excel,
    [Parameter(Mandatory=$true)][string]$FileName
  )
  $Workbook.SaveAs($FileName,52)
  if(-not (Test-Path -LiteralPath $FileName)) { throw 'XLSM was not created' }
  # Excel holds file handles even after SaveAs. Close BEFORE reading ZipArchive.
  $Workbook.Close($false)
  $Excel.Quit()
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  [GC]::Collect()
  Assert-TCXlsmArchive -FileName $FileName
}
function Assert-TCXlsmArchive {
  param([Parameter(Mandatory=$true)][string]$FileName)
  Add-Type -AssemblyName System.IO.Compression
  $stream=$null
  $lastError=''
  for($attempt=1;$attempt -le 12;$attempt++) {
    try {
      $stream=[System.IO.File]::Open($FileName,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
      break
    } catch [System.IO.IOException] {
      $lastError=$_.Exception.Message
      Start-Sleep -Milliseconds 450
    }
  }
  if($null -eq $stream) { throw ('XLSM created, but temporarily unavailable for inspection: '+$lastError) }
  try {
    $archive=New-Object System.IO.Compression.ZipArchive -ArgumentList @($stream,[System.IO.Compression.ZipArchiveMode]::Read)
    try {
      foreach($entryName in @('xl/vbaProject.bin','xl/workbook.xml','[Content_Types].xml')) {
        if($null -eq $archive.GetEntry($entryName)) { throw ('Required XLSM entry missing: '+$entryName) }
      }
    } finally { $archive.Dispose() }
  } finally { $stream.Dispose() }
  Write-Host ('XLSM package structure verified: '+$FileName)
}
