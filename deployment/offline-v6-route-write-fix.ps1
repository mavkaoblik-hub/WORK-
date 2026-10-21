# Offline v6 route import fix. Developer maintenance tool only.
# Applies the same scalar Excel COM write used in Transport_Center_Offline_v6.zip.
param([string]$InstallerPath = 'INSTALL_OFFLINE.ps1')
$ErrorActionPreference = 'Stop'
$path = [IO.Path]::GetFullPath($InstallerPath)
if (-not [IO.File]::Exists($path)) { throw ('Missing installer: ' + $path) }
$enc = New-Object System.Text.UTF8Encoding($true)
$s = [IO.File]::ReadAllText($path,[Text.Encoding]::UTF8)
$old = "  " + '$rt.Range(' + "'A6:M119'" + ').Value2=$data'
$new = @'
  # Excel COM can reject an Object[,] bulk assignment with Object[] -> String.
  # Write only scalar values to preserve number types and avoid COM array casts.
  $routeWritten=0
  for($ri=0;$ri -lt $items.Count;$ri++){
    $excelRow=$ri+6
    for($ci=0;$ci -lt 13;$ci++){
      $value=$data[$ri,$ci]
      if($null -ne $value -and -not ($value -is [DBNull])){
        if($value -is [System.Array]){throw "Invalid nested array in route $excelRow, col $ci"}
        $rt.Cells.Item($excelRow,$ci+1).Value2=$value
      }
    }
    $routeWritten++
  }
  Check ($routeWritten -eq 114) 'Not all 114 routes were entered.'
  Write-Host ('Route catalog: '+$routeWritten+' entries embedded.')
'@
if (-not $s.Contains($old)) { throw 'v5 bulk-write line not found; source may already be fixed.' }
$backup = $path + '.pre_v6'
if ([IO.File]::Exists($backup)) { throw 'Backup already exists; refusing overwrite.' }
[IO.File]::Copy($path,$backup)
[IO.File]::WriteAllText($path,$s.Replace($old,$new),$enc)
Write-Host ('Patched scalar route writes. Original saved as '+$backup)
