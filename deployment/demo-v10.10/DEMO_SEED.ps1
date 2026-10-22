#Requires -Version 5.1
param(
    [Parameter(Mandatory=$true)]$Book,
    [Parameter(Mandatory=$true)][string]$SeedPath
)
$ErrorActionPreference='Stop'
if(-not (Test-Path -LiteralPath $SeedPath)){throw "Missing DEMO fixture: $SeedPath"}
$demo=Get-Content -LiteralPath $SeedPath -Encoding UTF8 -Raw | ConvertFrom-Json
if($demo.kind -ne 'TRANSPORT_DEMO_SEED_V1' -or -not $demo.fictional){throw 'Refuse non-demo or unknown seed format'}
if(@($demo.drivers).Count -ne 7 -or @($demo.trips).Count -ne 75){throw 'Demo counts not as expected'}

function DemoText($s,[int]$row,[int]$col,$v){
    if($null -ne $v -and [string]$v -ne ''){$s.Cells.Item($row,$col).Value2=[string]$v}
}
function DemoNum($s,[int]$row,[int]$col,$v){
    if($null -ne $v -and [string]$v -ne ''){
        Set-ExcelNumber ($s.Cells.Item($row,$col)) ([double]$v) ("DEMO row $row col $col")
    }
}
function DemoDate($s,[int]$row,[int]$col,$v){
    if($null -ne $v -and [string]$v -ne ''){
        $text=[string]$v
        $text=$text.Replace('T',' ')
        $parsed=[datetime]::Parse($text,[System.Globalization.CultureInfo]::InvariantCulture)
        Set-ExcelNumber ($s.Cells.Item($row,$col)) $parsed.ToOADate() ("DEMO date row $row col $col")
    }
}

Write-Host 'DEMO: Adding 7 synthetic drivers and 9 vehicle profiles...' -ForegroundColor Cyan
$cars=$Book.Worksheets.Item('Автопарк')
for($i=0;$i -lt @($demo.vehicles).Count;$i++){
  $v=$demo.vehicles[$i];$r=6+$i
  DemoText $cars $r 1 $v.id;DemoText $cars $r 2 $v.brand
  DemoText $cars $r 3 $v.model;DemoText $cars $r 4 $v.plate
  DemoText $cars $r 5 $v.fuel
  DemoNum $cars $r 6 $v.consumption_l_100km
  DemoNum $cars $r 7 $v.price_pln_l
  DemoNum $cars $r 8 $v.tank_l
  DemoText $cars $r 9 'ДЕМО / активне'
  DemoText $cars $r 10 'ВИГАДАНІ ДАНІ. Тільки для тестування.'
}
$drivers=$Book.Worksheets.Item('Контакти_водіїв')
$registry=$Book.Worksheets.Item('Реєстр_водіїв')
$template=$Book.Worksheets.Item('ШАБЛОН_ВОДІЯ')
$endMark=$Book.Worksheets.Item('КІНЕЦЬ_ВОДІЇВ')
$templateState=$template.Visible
try {
 $template.Visible=-1
 for($i=0;$i -lt @($demo.drivers).Count;$i++){
   $d=$demo.drivers[$i];$r=6+$i
   DemoText $drivers $r 1 $d.id
   DemoText $drivers $r 2 $d.surname
   DemoText $drivers $r 3 $d.name
   DemoText $drivers $r 5 $d.car
   DemoText $drivers $r 7 'Активний / ДЕМО'
   DemoText $registry (13+$i) 1 $d.id
   if($i -ge 2){
      # Copy hidden driver template into the existing driver section; normal
      # editing/driver macros expect these worksheets to exist.
      $template.Copy($endMark)
      $newDriverSheet=$Book.Application.ActiveSheet
      $newDriverSheet.Name=[string]$d.id
   }else{
      $newDriverSheet=$Book.Worksheets.Item([string]$d.id)
   }
   DemoText $newDriverSheet 5 2 ($d.surname+' '+$d.name)
   DemoText $newDriverSheet 5 5 $demo.vehicles[$i].plate
   $newDriverSheet.Visible=0
 }
} finally {$template.Visible=$templateState}

# Clearly marked, non-market FX examples: PLN per 1 currency unit.
$fx=$Book.Worksheets.Item('Курси валют')
DemoNum $fx 7 2 $demo.fx_pln_per_unit.USD
DemoNum $fx 8 2 $demo.fx_pln_per_unit.EUR
DemoNum $fx 9 2 $demo.fx_pln_per_unit.UAH
foreach($r in 7..9){
 DemoDate $fx $r 3 $demo.as_of
 DemoText $fx $r 4 'ДЕМО-КУРС / не NBP / не використовувати в обліку'
 DemoDate $fx $r 5 $demo.as_of
}
$fx.Range('A23').Value2='УВАГА: УСІ КУРСИ ДЕМОНСТРАЦІЙНІ, НЕ РЕАЛЬНІ.'
$fx.Range('A23').Font.Bold=$true
$fx.Range('A23').Font.Color=255
$fx.Range('A24').Value2='Для справжніх курсів натисніть Оновити NBP або внесіть власні.'

Write-Host 'DEMO: Seeding completed, active and planned trips...' -ForegroundColor Cyan
$trips=$Book.Worksheets.Item('TC_Рейси')
$tripIndex=0
foreach($t in $demo.trips){
  $r=6+$tripIndex
  $tripIndex++
  DemoText $trips $r 1 $t.id
  DemoDate $trips $r 2 $t.created
  DemoText $trips $r 3 $t.driver
  DemoText $trips $r 4 $t.car
  DemoText $trips $r 5 $t.route
  DemoText $trips $r 6 $t.customer
  DemoDate $trips $r 7 $t.plan_start
  DemoText $trips $r 8 $t.status
  DemoDate $trips $r 9 $t.actual_start
  DemoDate $trips $r 10 $t.actual_end
  DemoNum $trips $r 11 $t.odometer_start
  DemoNum $trips $r 12 $t.odometer_end
  DemoNum $trips $r 13 $t.km_actual
  DemoNum $trips $r 14 $t.revenue_usd
  DemoNum $trips $r 15 $t.paid_usd
  DemoNum $trips $r 16 $t.km_plan
  DemoNum $trips $r 17 $t.fuel_liters_plan
  DemoNum $trips $r 18 $t.fuel_cost_pln_plan
  # Write dynamic ledger formulas exactly as the production SaveNew procedure.
  $trips.Cells.Item($r,19).Formula="=SUMIF('TC_Витрати'!`$B`$6:`$B`$50000,A$r,'TC_Витрати'!`$H`$6:`$H`$50000)"
  $trips.Cells.Item($r,20).Formula="=N$r-S$r"
  DemoText $trips $r 21 $t.memo
  DemoDate $trips $r 22 $t.finished_at
  DemoNum $trips $r 23 $t.driving_h
  DemoNum $trips $r 24 $t.waiting_h
  DemoNum $trips $r 25 $t.lost_h
  DemoDate $trips $r 26 $t.plan_end
  DemoNum $trips $r 27 $t.odometer_now
  DemoDate $trips $r 28 $t.odometer_updated
}
Write-Host 'DEMO: Seeding categories, receipts and 4-currency expense history...' -ForegroundColor Cyan
$spend=$Book.Worksheets.Item('TC_Витрати')
$idx=0
foreach($x in $demo.expenses){
 $r=6+$idx
 DemoDate $spend $r 1 $x.created
 DemoText $spend $r 2 $x.trip
 DemoText $spend $r 3 $x.category
 DemoText $spend $r 4 $x.parking
 DemoNum $spend $r 5 $x.amount
 DemoText $spend $r 6 $x.currency
 DemoNum $spend $r 7 $x.fx_pln_per_unit
 DemoNum $spend $r 8 $x.usd
 DemoDate $spend $r 9 $x.date
 DemoText $spend $r 10 $x.note
 DemoText $spend $r 11 'DEMO'
 $idx++
}
# Date serials are numeric (never text), but formatting should carry Date types
# through Excel .Value for VBA routines that use IsDate / DateValue.
Set-NumberFormatSafe ($trips.Range('B6:B80')) 'dd.mm.yyyy hh:mm' 'DEMO trip created'
foreach($c in @('G','I','J','V','Z','AB')){
 Set-NumberFormatSafe ($trips.Range("$($c)6:$($c)80")) 'dd.mm.yyyy' ('DEMO trip date '+$c)
}
Set-NumberFormatSafe ($spend.Range('A6:A270')) 'dd.mm.yyyy hh:mm' 'DEMO expense created'
Set-NumberFormatSafe ($spend.Range('I6:I270')) 'dd.mm.yyyy' 'DEMO spending date'
if($tripIndex -ne 75 -or $idx -ne 259){throw ("Incomplete DEMO seed: trips=$tripIndex, expenses=$idx")}
$Book.Worksheets.Item('Зміст').Range('B38').Value2='ДЕМО! 7 водіїв, 9 авто, 75 рейсів, 259 витрат. Усі дані вигадані.'
$Book.Worksheets.Item('Зміст').Range('B38').Font.Bold=$true
$Book.Worksheets.Item('Зміст').Range('B38').Font.Color=65535
Write-Host 'DEMO fixture loaded: 7 drivers, 9 vehicles, 75 trips and 259 expenses.' -ForegroundColor Green