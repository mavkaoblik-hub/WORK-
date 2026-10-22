#Requires -Version 5.1
[CmdletBinding()]
param([Parameter(Mandatory=$true)]$Sheet)
$ErrorActionPreference='Stop'
function RGB([int]$r,[int]$g,[int]$b) { return ($r+256*$g+65536*$b) }
$ws=$Sheet
# Design-only: no source journal values are changed.
$ws.Range('A1:R80').UnMerge()
$ws.Range('A1:R80').ClearContents()
$canvas=$ws.Range('A1:R80')
$canvas.Interior.Color=(RGB 245 248 250)
$canvas.Font.Name='Segoe UI'
$canvas.Font.Size=11
$canvas.Font.Color=(RGB 35 53 59)
$canvas.RowHeight=22
foreach($col in @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R')) {
  $ws.Columns.Item($col).ColumnWidth=10.3
}
$ws.Columns.Item('A').ColumnWidth=3
$ws.Columns.Item('R').ColumnWidth=3
$ws.Rows.Item(1).RowHeight=9
$ws.Range('A1:R1').Interior.Color=(RGB 53 132 111)
$ws.Rows.Item(2).RowHeight=35
$ws.Rows.Item(3).RowHeight=28
$ws.Rows.Item(4).RowHeight=25
$ws.Range('B2:Q4').Interior.Color=(RGB 255 255 255)
$ws.Range('B2:Q3').Merge()
$ws.Range('B2').Value2='ТРАНСПОРТНИЙ ЦЕНТР'
$ws.Range('B2').Font.Bold=$true
$ws.Range('B2').Font.Size=23
$ws.Range('B2').Font.Color=(RGB 30 82 69)
$ws.Range('B2:Q3').VerticalAlignment=-4108
$ws.Range('B4:Q4').Merge()
$ws.Range('B4').Value2='Робочий кабінет   /   рейси, автопарк, витрати й аналітика'
$ws.Range('B4').Font.Color=(RGB 107 123 129)
$ws.Range('B4').Font.Size=11
$ws.Rows.Item(5).RowHeight=17
$ws.Range('B6:Q6').Merge()
$ws.Range('B6').Value2='ОГЛЯД ДІЯЛЬНОСТІ'
$ws.Range('B6').Font.Bold=$true
$ws.Range('B6').Font.Size=12
$ws.Range('B6').Font.Color=(RGB 42 92 82)
foreach($row in @(7,11)){$ws.Rows.Item($row).RowHeight=25}
foreach($row in @(8,12)){$ws.Rows.Item($row).RowHeight=26}
foreach($row in @(9,13)){$ws.Rows.Item($row).RowHeight=24}
$cards = @(
 @('B7:D9','B7:D7','B8:D9','B7','ВОДІЇ','B8','=COUNTA(Контакти_водіїв!$A$6:$A$1000)'),
 @('F7:H9','F7:H7','F8:H9','F7','ЗАПЛАНОВАНО','F8','=COUNTIF(~TC_Рейси~!$H$6:$H$50000,"Заплановано")'),
 @('J7:L9','J7:L7','J8:L9','J7','У РЕЙСІ','J8','=COUNTIF(~TC_Рейси~!$H$6:$H$50000,"У рейсі")'),
 @('N7:P9','N7:P7','N8:P9','N7','ЗАВЕРШЕНО','N8','=COUNTIF(~TC_Рейси~!$H$6:$H$50000,"Завершено")'),
 @('B11:D13','B11:D11','B12:D13','B11','МАРШРУТІВ','B12','=COUNTA(Маршрути_Європа!$A$6:$A$1000)'),
 @('F11:H13','F11:H11','F12:H13','F11','ПАРКІНГІВ','F12','=COUNTA(~TC_Паркінги~!$A$6:$A$1000)'),
 @('J11:L13','J11:L11','J12:L13','J11','ВИТРАТИ · USD','J12','=SUM(~TC_Витрати~!$H$6:$H$50000)'),
 @('N11:P13','N11:P11','N12:P13','N11','ДОХІД · USD','N12','=SUM(~TC_Рейси~!$N$6:$N$50000)')
)
foreach($item in $cards){
  $whole=$ws.Range([string]$item[0])
  $header=$ws.Range([string]$item[1])
  $value=$ws.Range([string]$item[2])
  $whole.Interior.Color=(RGB 255 255 255)
  try {$whole.Borders.LineStyle=1;$whole.Borders.Color=(RGB 221 233 229)}catch{}
  $header.Merge()
  $value.Merge()
  $ws.Range([string]$item[3]).Value2=[string]$item[4]
  $ws.Range([string]$item[3]).Font.Size=11
  $ws.Range([string]$item[3]).Font.Color=(RGB 104 122 126)
  $ws.Range([string]$item[3]).Font.Bold=$true
  $ws.Range([string]$item[5]).Formula=([string]$item[6]).Replace('~',[string][char]39)
  $ws.Range([string]$item[5]).Font.Size=21
  $ws.Range([string]$item[5]).Font.Bold=$true
  $ws.Range([string]$item[5]).Font.Color=(RGB 26 87 72)
  $header.HorizontalAlignment=-4108
  $value.HorizontalAlignment=-4108
  $header.VerticalAlignment=-4108
  $value.VerticalAlignment=-4108
}
foreach($cell in @('J12','N12')){
  try {$ws.Range($cell).NumberFormat='#,##0.00'}catch{}
}
$groupHeaders=@(
 @('B15:Q15','ОПЕРАЦІЇ З РЕЙСАМИ'),
 @('B20:Q20','ДОВІДНИКИ ТА АВТОПАРК'),
 @('B25:Q25','ФІНАНСИ ТА ЗВІТНІСТЬ')
)
foreach($line in $groupHeaders){
  $r=$ws.Range([string]$line[0])
  $r.Merge()
  $r.Value2=[string]$line[1]
  $r.Interior.Color=(RGB 226 240 234)
  $r.Font.Color=(RGB 30 87 73)
  $r.Font.Size=12
  $r.Font.Bold=$true
  $r.RowHeight=28
  $r.VerticalAlignment=-4108
}
foreach($r in @(17,22,27,32)){$ws.Rows.Item($r).RowHeight=45}
$ws.Range('B36:Q37').Merge()
$ws.Range('B36').Value2='Дані зберігаються у внутрішніх журналах. Панель відображає облікові показники, а дії відкривають окремі форми.'
$ws.Range('B36:Q37').Interior.Color=(RGB 235 245 240)
$ws.Range('B36').Font.Color=(RGB 86 114 107)
$ws.Range('B36').Font.Size=10
$ws.Range('B36:Q37').WrapText=$true
$ws.Range('B36:Q37').VerticalAlignment=-4108
$ws.Range('B40:Q40').Merge()
$ws.Range('B40').Value2='CLEAN · v10.15.1   |   Без демонстраційних даних'
$ws.Range('B40').Font.Color=(RGB 137 154 158)
$ws.Range('B40').Font.Size=10
Write-Host 'v10.15.1: light dashboard polished; ledgers unchanged.' -ForegroundColor Green
