# This helper is mirrored in the customer-supplied offline INSTALL_OFFLINE.ps1 (v7).
# Numeric payloads are injected into Excel through Formula (a string COM argument),
# rather than using late-bound Value2 = Double that fails on some Excel LTSC setups.
function Set-ExcelNumber($cell, [double]$number, [string]$cellInfo) {
  if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) {throw ("Non-finite numeric value for " + $cellInfo)}
  $numText = $number.ToString('0.##############',[System.Globalization.CultureInfo]::InvariantCulture)
  try { $cell.Formula = ('=' + $numText) }
  catch { throw ("Excel numeric write failed for " + $cellInfo + ' (' + $numText + '): ' + $_.Exception.Message) }
}
function Test-NumberSetter {
  foreach($case in @(
     @{num=594.0;expected='=594'},
     @{num=5.25;expected='=5.25'},
     @{num=0.1;expected='=0.1'},
     @{num=2018.0;expected='=2018'}
  )) {
    # A string-typed property rejects numeric Double assignment and verifies
    # that the helper passes an invariant text formula instead.
    $mock=[pscustomobject]@{Formula=[string]''}
    Set-ExcelNumber $mock ([double]$case.num) 'mock'
    if($mock.Formula -cne $case.expected) {throw ("Expected "+$case.expected+", received "+$mock.Formula)}
  }
  $failed=$false
  try { Set-ExcelNumber ([pscustomobject]@{Formula=''}) ([double]::NaN) 'bad' }
  catch { $failed=$true }
  if(-not $failed) {throw 'Non-finite number should be rejected'}
  Write-Host 'OFFLINE v7 numeric formula tests passed'
}
if($MyInvocation.InvocationName -ne '.') { Test-NumberSetter }
