#Requires -Version 5.1
param()
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $dir 'Транспортний_центр_v8_робочий_Times_New_Roman.xlsx'
$target = Join-Path $dir 'Транспортний_центр.xlsm'
$moduleFile = Join-Path $dir 'vba\TransportCenter.bas'
$appModuleFile = Join-Path $dir 'vba\TransportApp.bas'
$insightsFile = Join-Path $dir 'vba\TransportInsights.bas'
$statusFxFile = Join-Path $dir 'vba\TransportStatusFX.bas'
$routesFile = Join-Path $dir 'data\routes_europe.json'
$excel = $null; $book = $null; $guard = $null
function Check([bool]$ok,[string]$message){if(-not $ok){throw $message}}
function Button($ws,[string]$label,[string]$proc,[double]$x,[double]$y){
  $s = $ws.Shapes.AddShape(5,$x,$y,185,34)
  $s.TextFrame.Characters().Text = $label
  $s.TextFrame.Characters().Font.Name = 'Times New Roman'
  $s.TextFrame.Characters().Font.Size = 13
  $s.OnAction = "'Транспортний_центр.xlsm'!TransportCenter.$proc"
  $s.Placement=3
  try{
    $s.Fill.ForeColor.RGB=(37+80*256+69*65536) # forest-green buttons
    $s.Line.Visible=$false
    $s.TextFrame.Characters().Font.Color=16777215
    $s.TextFrame.Characters().Font.Bold=$true
    $s.TextFrame.Characters().ParagraphFormat.Alignment=2
    $s.Shadow.Visible=$false
  }catch{}
}
# Anchor buttons to worksheet cells instead of hard-coded coordinates.
# Header rows and control rows stay disjoint at every Excel zoom level.
function DashboardButton($ws,[string]$label,[string]$proc,[string]$anchor){
  $cell=$ws.Range($anchor)
  Button $ws $label $proc ([double]$cell.Left) ([double]$cell.Top)
}
# Microsoft Excel LTSC COM can reject Range.NumberFormat with 0x800A03EC.
# Cell formatting is presentation only, NEVER abort the entire workbook build over it.
# Preserve numeric/date VALUES and formulas; use NumberFormatLocal if standard fails.
function Set-NumberFormatSafe($range, [string]$pattern, [string]$address) {
    try {
        $range.NumberFormat = $pattern
        return
    } catch {
        try {
            $range.NumberFormatLocal = $pattern
            return
        } catch {
            Write-Warning ("Excel refused cosmetic number format for " + $address + "; default cell format preserved.")
        }
    }
}
# Numbers are written as invariant numeric literals through Excel's Formula property.
# COM Formula takes a String; Excel then evaluates a true numeric value (not a text cell).
# Avoid direct Value2 = Double, rejected by some PowerShell 5.1 + Excel LTSC setups.
function Set-ExcelNumber($cell, [double]$number, [string]$cellInfo) {
  if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) {throw ("Non-finite numeric value for " + $cellInfo)}
  $numText = $number.ToString('0.##############',[System.Globalization.CultureInfo]::InvariantCulture)
  try { $cell.Formula = ('=' + $numText) }
  catch { throw ("Excel numeric write failed for " + $cellInfo + ' (' + $numText + '): ' + $_.Exception.Message) }
}
function Set-ExcelRouteCell($sheet,[int]$excelRow,[int]$excelColumn,$item,[string]$routeId) {
  try {
    $cell = $sheet.Cells.Item($excelRow,$excelColumn)
    if($item -is [int] -or $item -is [long] -or $item -is [float] -or $item -is [double] -or $item -is [decimal]) {
      Set-ExcelNumber $cell ([double]$item) ("Route " + $routeId + ', row ' + $excelRow + ', col ' + $excelColumn)
    } else {
      $cell.Value2 = [string]$item
    }
  } catch { throw ("Failed route " + $routeId + ', Excel row ' + $excelRow + ', col ' + $excelColumn + ': ' + $_.Exception.Message) }
}
# Avoid nested PowerShell apostrophe escaping inside Excel formulas.
# A tilde in formula templates stands for Excel's single quote around a sheet name.
function ExcelExpr([string]$template){
  return $template.Replace('~',[string][char]39)
}
function Catalog($book,[string]$name,[string[]]$headers,[string]$last,[int]$lastRow){
  $ws = $book.Worksheets.Add()
  $ws.Name = $name
  $ws.Range('A1').Value2 = "ТРАНСПОРТНИЙ ЦЕНТР / $name"
  $ws.Range('A1').Font.Bold = $true
  $ws.Range('A1').Font.Size = 18
  $ws.Range('A3').Value2 = 'Жовті поля для ручного внесення. Інтернет-оновлення автоматично не виконується.'
  for($i=0; $i -lt $headers.Count; $i++){$ws.Cells.Item(5,$i+1).Value2=$headers[$i]}
  $ws.Range("A5:$($last)5").Font.Bold = $true
  $ws.Range("A5:$($last)5").Interior.Color = 2316839
  $ws.Range("A5:$($last)5").Font.Color = 16777215
  $ws.Range("A1:$($last)$lastRow").Font.Name = 'Times New Roman'
  $ws.Range("A1:$($last)$lastRow").Font.Size = 14
  $ws.Rows.Item(5).RowHeight=37
  $ws.Rows.Item(5).WrapText=$true
  $ws.Columns.Item(1).ColumnWidth=23
  return $ws
}
try {
  Check (Test-Path -LiteralPath $src) 'Не знайдено початкової Excel-книги поруч із інсталятором.'
  Check (Test-Path -LiteralPath $moduleFile) 'Немає локального VBA-модуля.'
  Check (Test-Path -LiteralPath $appModuleFile) 'Немає модуля інтерфейсу v10.'
  Check (Test-Path -LiteralPath $insightsFile) 'Немає VBA-модуля звітності та стану.'
  Check (Test-Path -LiteralPath $statusFxFile) 'Немає модуля поточного стану, архіву й курсів валют.'
  Check (Test-Path -LiteralPath (Join-Path $dir 'vba\TransportWorkbookEvents.txt')) 'Немає коду автоматичного оновлення аркушів.'
  Check (Test-Path -LiteralPath $routesFile) 'Немає локального довідника маршрутів.'
  Check (-not (Test-Path -LiteralPath $target)) 'Транспортний_центр.xlsm вже існує. Автоматичний перезапис заборонено.'
  Write-Host '1/5. Запуск Microsoft Excel без мережевих запитів...' -ForegroundColor Cyan
  $excel=New-Object -ComObject Excel.Application
  $excel.Visible=$false
  $excel.DisplayAlerts=$false
  $excel.ScreenUpdating=$false
  $excel.EnableEvents=$false
  $excel.AskToUpdateLinks=$false
  $excel.AutomationSecurity=3
  # On Excel LTSC, setting Application.Calculation before a book exists may fail 0x800A03EC.
  # Open a disposable blank book FIRST, then disable calc BEFORE opening historical web formulas.
  $guard=$excel.Workbooks.Add()
  $excel.Calculation=-4135
  # Fail FAST before editing the source workbook when VBA programmatic access
  # is disabled by Trust Center / Group Policy or VBA itself is not installed.
  Write-Host '1/5. Перевірка дозволу Excel на програмне створення VBA...' -ForegroundColor Cyan
  $probeComponent = $null
  $probeForm = $null
  try {
    $probeProject = $guard.VBProject
    if ($null -eq $probeProject) { throw 'VBProject повернув порожній об`єкт (null).' }
    $probeComponents = $probeProject.VBComponents
    if ($null -eq $probeComponents) { throw 'VBComponents повернув порожній об`єкт (null).' }
    $probeComponent = $probeComponents.Add(1)
    if ($null -eq $probeComponent) { throw 'Excel не створив тестовий VBA-модуль.' }
    $probeComponent.Name = 'TC_AccessPreflight'
    $probeCode = $probeComponent.CodeModule
    if ($null -eq $probeCode) { throw 'CodeModule повернув порожній об`єкт (null).' }
    $probeCode.AddFromString("Option Explicit`r`nPublic Sub TC_AccessPreflight()`r`nEnd Sub")
    # Verify MSForms design-time object model BEFORE building actual catalogues.
    # Some Excel COM versions return Designer objects without Width/Height.
    $probeForm = $probeComponents.Add(3)
    if($null -eq $probeForm) { throw 'Microsoft Forms UserForm не створено.' }
    $probeForm.Name = 'TC_FormPreflight'
    $probeForm.Properties.Item('Caption').Value = 'TC Form Preflight'
    $probeForm.Properties.Item('Width').Value = 210
    $probeForm.Properties.Item('Height').Value = 105
    $probeDesigner = $probeForm.Designer
    if($null -eq $probeDesigner) { throw 'Не отримано Designer тестової форми.' }
    $probeControl = $probeDesigner.Controls.Add('Forms.CommandButton.1','cmdPreflight',$true)
    if($null -eq $probeControl) { throw 'MSForms.Controls.Add не створив CommandButton.' }
    $probeControl.Move([single]12,[single]12,[single]120,[single]25)
    $probeControl.Caption = 'OK'
    Write-Host 'Доступ до VBA та створення UserForm перевірено успішно.' -ForegroundColor Green
  } catch {
    $details = $_.Exception.Message
    throw ('VBA_ACCESS_BLOCKED: Excel заборонив програмне вбудовування макросів. ' +
           'У Microsoft Excel відкрийте Файл > Параметри > Центр безпеки > Параметри центру безпеки > Параметри макросів, ' +
           'увімкніть Довіряти доступу до об`єктної моделі проєктів VBA та натисніть OK. ' +
           'Закрийте вікно Excel з налаштуваннями та повторно запустіть інсталятор. ' +
           'Після успішного створення XLSM вимкніть цей тимчасовий дозвіл. ' +
           'Якщо перемикач заблокований адміністратором або VBA-компонент не встановлений, зверніться до адміністратора Office. ' +
           'Технічні деталі: ' + $details)
  } finally {
    if ($null -ne $probeForm) {
      try { $probeComponents.Remove($probeForm) } catch { Write-Warning 'Не вдалося видалити тимчасову UserForm.' }
    }
    if ($null -ne $probeComponent) {
      try { $probeComponents.Remove($probeComponent) } catch { Write-Warning 'Не вдалося видалити тимчасовий тестовий модуль.' }
    }
    $probeForm = $null
    $probeComponent = $null
    $probeProject = $null
    $probeComponents = $null
  }
  $book=$excel.Workbooks.Open($src,0,$false)
  $guard.Close($false)
  $guard=$null
  # Remove all previous WEBSERVICE / FILTERXML formulas before enabling Excel recalculation.
  $book.Worksheets.Item('NBP_дані').Range('B5:B8').ClearContents()
  $book.Worksheets.Item('Курси_NBP').Range('G13:H52').ClearContents()
  $book.Worksheets.Item('Налаштування').Range('B11').Value2='Курс NBP тільки після кнопки'
  $book.Worksheets.Item('Налаштування').Range('B15').Value2='Фонових інтернет-оновлень немає'
  $book.Worksheets.Item('Інструкція').Range('B7').Value2='Інтернет лише вручну: кнопка Оновити NBP. GitHub ніяк не використовується.'
  Write-Host '2/5. Вбудовані каталоги автопарку, водіїв і 114 маршрутів...'
  $car=Catalog $book 'Автопарк' @('Код','Марка','Модель','Номер','Пальне','Норма л/100км','PLN/літр','Бак л','Статус','Примітка') 'J' 65
  foreach($col in @('A','B','C','D','E','F','G','H','I','J')){$car.Columns.Item($col).ColumnWidth=22}
  $car.Columns.Item('J').ColumnWidth=38
  $car.Range('A6:J65').Interior.Color=14348287
  Set-NumberFormatSafe ($car.Range('F6:H65')) '0.00' 'F6:H65'
  $book.Names.Add('TC_Cars','=Автопарк!$A$6:$A$65')|Out-Null
  $cnt=Catalog $book 'Контакти_водіїв' @('Аркуш','Прізвище','Імя','Телефон','Авто ID','Email','Статус','Вибір у формі') 'H' 65
  foreach($col in @('A','B','C','D','E','F','G','H')){$cnt.Columns.Item($col).ColumnWidth=23}
  $cnt.Columns.Item('F').ColumnWidth=33
  $cnt.Columns.Item('H').ColumnWidth=38
  Set-NumberFormatSafe ($cnt.Range('D6:D65')) '@' 'D6:D65'
  $cnt.Range('A6:G65').Interior.Color=14348287
  $cnt.Range('H6:H65').Interior.Color=14277081
  # CLEAN: driver IDs are created only when a user registers a driver.
  $cnt.Range('H6').Formula='=IF(A6="","",IF(AND(B6="",C6=""),A6,TRIM(B6&" "&C6)&" ["&A6&"]"))'
  $cnt.Range('H6:H65').FillDown()
  $book.Names.Add('TC_Drivers','=Контакти_водіїв!$H$6:$H$65')|Out-Null
  $cnt.Range('E6:E65').Validation.Add(3,1,1,'=TC_Cars')
  $rt=Catalog $book 'Маршрути_Європа' @('Код','Звідки','Куди','Км','Години','Метод','Рейсів','Км факт','Закрито','Статус','URL джерела','Рік','Примітка') 'M' 205
  $rt.Range('A3').Value2='Відстань із публікацій. 56 зворотних значень є лише оцінкою. Перед вантажним рейсом перевірте маршрут.'
  foreach($col in @('A','B','C','D','E','F','G','H','I','J','K','L','M')){$rt.Columns.Item($col).ColumnWidth=21}
  $rt.Columns.Item('F').ColumnWidth=38
  $rt.Columns.Item('J').ColumnWidth=26
  $rt.Columns.Item('K').ColumnWidth=55
  $rt.Columns.Item('M').ColumnWidth=60
  $raw=Get-Content -LiteralPath $routesFile -Raw -Encoding UTF8 | ConvertFrom-Json
  $items=@($raw)
  Check ($items.Count -eq 114) "Очікувалось 114 маршрутів; знайдено $($items.Count)."
  $data=[object[,]]::new($items.Count,13)
  $unique=@{}
  for($i=0;$i -lt $items.Count;$i++){
    $v=$items[$i]; $code=[string]$v.id
    Check (-not $unique.ContainsKey($code)) "Повтор маршруту: $code"
    Check ([double]$v.km -gt 0) "Немає кілометражу: $code"
    $unique[$code]=$true
    $data[$i,0]=$code
    $data[$i,1]=[string]$v.start
    $data[$i,2]=[string]$v.end
    $data[$i,3]=[double]$v.km
    if($null -ne $v.hours){$data[$i,4]=[double]$v.hours}
    $data[$i,5]=[string]$v.typ
    $data[$i,10]=[string]$v.source
    $data[$i,11]=[int]$v.year
    $data[$i,12]=[string]$v.note
  }
  # PowerShell 5.1 + Excel COM can marshal Object[,] as Object[] and then
  # attempt to cast the array into String. This is fatal (Excel LTSC).
  # Write only occupied scalars, one cell at a time; 114 routes is small.
  # Keep values numeric where numeric; do not depend on regional separators.
  $routeWritten=0
  for($ri=0;$ri -lt $items.Count;$ri++){
    $excelRow=$ri+6
    $routeCode=[string]$items[$ri].id
    for($ci=0;$ci -lt 13;$ci++){
      $value=$data[$ri,$ci]
      if($null -ne $value -and -not ($value -is [DBNull])){
        if($value -is [System.Array]){throw "Invalid nested array in route $excelRow, col $ci"}
        Set-ExcelRouteCell $rt $excelRow ($ci+1) $value $routeCode
      }
    }
    $routeWritten++
  }
  Check ($routeWritten -eq 114) 'Не всі маршрути записано в Excel.'
  Write-Host ('Каталог: '+$routeWritten+' маршрутів вбудовано у книгу.')
  $rt.Range('G6').Formula='=IF(A6="","",COUNTIF(Нарахування!$O$13:$O$132,A6))'
  $rt.Range('H6').Formula='=IF(A6="","",SUMIF(Нарахування!$O$13:$O$132,A6,Нарахування!$R$13:$R$132))'
  $rt.Range('I6').Formula='=IF(A6="","",COUNTIFS(Нарахування!$O$13:$O$132,A6,Нарахування!$L$13:$L$132,"ЗАКРИТО"))'
  $rt.Range('J6').Formula='=IF(A6="","",IF(ISNUMBER(SEARCH("Зворотн",F6)),"ЗВОРОТНИЙ ОРІЄНТИР",IF(L6<2024,"СТАРІ ДАНІ","ПЕРЕВІРИТИ ДЛЯ TIR")))'
  foreach($letter in @('G','H','I','J')){$rt.Range("$($letter)6:$($letter)205").FillDown()}
  $rt.Range('G6:J205').Interior.Color=14277081
  $book.Names.Add('TC_Routes','=Маршрути_Європа!$A$6:$A$205')|Out-Null
  Write-Host '3/5. Адміністративна форма та фінансові розрахунки...'
  $f=$book.Worksheets.Add();$f.Name='Форма_адміністратора'
  $f.Range('A1').Value2='ТРАНСПОРТНИЙ ЦЕНТР / ОФОРМЛЕННЯ РЕЙСУ'
  $f.Range('A1').Font.Bold=$true;$f.Range('A1').Font.Size=18
  $f.Range('A3').Value2='Жовті поля заповнює адміністратор. Сині розраховуються автоматично.'
  $inputs=@('Водій','Дата нарахування','Замовник','Маршрут ID','Виїзд','Прибуття','Одометр поч.','Одометр кін.','Дохід USD','Отримано USD','Кермування, год','Простій, год','Примітка','Втрати часу','Авто ID (заміна)','Пальне PLN/л (заміна)','Пробіг, км (заміна)','Резерв пального, %')
  for($i=0;$i -lt $inputs.Count;$i++){$f.Cells.Item($i+5,1).Value2=$inputs[$i]}
  $results=@('Водій','Авто ID','Марка / номер','Норма л/100 км','Км план','PLN/л','Прогноз літрів','Прогноз PLN','Прогноз USD','Стан','Всього рейсів','Маршрут')
  for($i=0;$i -lt $results.Count;$i++){$f.Cells.Item($i+7,5).Value2=$results[$i]}
  $f.Range('E5').Value2='ПРОГНОЗ ПАЛЬНОГО'
  $f.Columns.Item('A').ColumnWidth=35
  $f.Columns.Item('B').ColumnWidth=38
  $f.Columns.Item('E').ColumnWidth=29
  $f.Columns.Item('F').ColumnWidth=45
  $f.Range('A1:F23').Font.Name='Times New Roman'
  $f.Range('A1:F23').Font.Size=14
  $f.Range('B5:B22').Interior.Color=14348287
  $f.Range('F7:F18').Interior.Color=14277081
  # COM-safe numeric Excel date serials. TC_ShowForm rewrites with today's date each time.
  $todaySerial = [double]([datetime]::Today.ToOADate())
  Set-ExcelNumber ($f.Range('B6')) $todaySerial 'Form B6: issue date'
  Set-ExcelNumber ($f.Range('B9')) $todaySerial 'Form B9: start date'
  Set-ExcelNumber ($f.Range('B10')) $todaySerial 'Form B10: arrival date'
  Set-NumberFormatSafe ($f.Range('B6')) 'dd.mm.yyyy' 'B6'
  Set-NumberFormatSafe ($f.Range('B9:B10')) 'dd.mm.yyyy' 'B9:B10'
  Set-ExcelNumber ($f.Range('B22')) 0.1 'Form B22: fuel reserve'
  Set-NumberFormatSafe ($f.Range('B22')) '0%' 'B22'
  $f.Range('B5').Validation.Add(3,1,1,'=TC_Drivers')
  $f.Range('B8').Validation.Add(3,1,1,'=TC_Routes')
  $f.Range('B19').Validation.Add(3,1,1,'=TC_Cars')
  $formulas=@{
  'F7'='=IF(B5="","",B5)';
  'F8'='=IF(B19<>"",B19,IFERROR(INDEX(Контакти_водіїв!$E$6:$E$65,MATCH(B5,Контакти_водіїв!$H$6:$H$65,0)),""))';
  'F9'='=IF(F8="","",IFERROR(INDEX(Автопарк!$B$6:$B$65,MATCH(F8,Автопарк!$A$6:$A$65,0))&" / "&INDEX(Автопарк!$D$6:$D$65,MATCH(F8,Автопарк!$A$6:$A$65,0)),""))';
  'F10'='=IF(F8="","",IFERROR(INDEX(Автопарк!$F$6:$F$65,MATCH(F8,Автопарк!$A$6:$A$65,0)),0))';
  'F11'='=IF(B21>0,B21,IFERROR(INDEX(Маршрути_Європа!$D$6:$D$205,MATCH(B8,Маршрути_Європа!$A$6:$A$205,0)),0))';
  'F12'='=IF(B20>0,B20,IFERROR(INDEX(Автопарк!$G$6:$G$65,MATCH(F8,Автопарк!$A$6:$A$65,0)),0))';
  'F13'='=IF(OR(F10<=0,F11<=0),"",ROUND(F10*F11/100,2))';
  'F14'='=IF(OR(F13="",F12<=0),"",ROUND(F13*F12*(1+B22),2))';
  'F15'='=IF(OR(F14="",Налаштування!$B$8<=0),"",ROUND(F14/Налаштування!$B$8,2))';
  'F16'='=IF(AND(F8<>"",F10>0,F11>0,F12>0),"ГОТОВО","УТОЧНИТИ АВТО / КМ / ЦІНУ")';
  'F17'='=IF(B8="","",COUNTIF(Нарахування!$O$13:$O$132,B8))';
  'F18'='=IF(B8="","",IFERROR(INDEX(Маршрути_Європа!$B$6:$B$205,MATCH(B8,Маршрути_Європа!$A$6:$A$205,0))&" — "&INDEX(Маршрути_Європа!$C$6:$C$205,MATCH(B8,Маршрути_Європа!$A$6:$A$205,0)),""))'
  }
  foreach($key in $formulas.Keys){$f.Range($key).Formula=[string]$formulas[$key]}
  Set-NumberFormatSafe ($f.Range('F10:F15')) '0.00' 'F10:F15'
  $card=$book.Worksheets.Add();$card.Name='Картка_водія'
  $card.Range('A1').Value2='ТРАНСПОРТНИЙ ЦЕНТР / ДОДАТИ ВОДІЯ'
  $card.Range('A1').Font.Bold=$true
  $card.Range('A1').Font.Size=18
  $profile=@('Прізвище','Імя','Телефон','E-mail','Авто ID','Назва вкладки (необовязково)')
  for($i=0;$i -lt $profile.Count;$i++){$card.Cells.Item($i+6,1).Value2=$profile[$i]}
  $card.Columns.Item('A').ColumnWidth=39
  $card.Columns.Item('B').ColumnWidth=40
  $card.Range('A1:B12').Font.Name='Times New Roman'
  $card.Range('A1:B12').Font.Size=14
  $card.Range('B6:B11').Interior.Color=14348287
  Set-NumberFormatSafe ($card.Range('B8')) '@' 'B8'
  $card.Range('B10').Validation.Add(3,1,1,'=TC_Cars')
  $audit=Catalog $book 'Журнал_операцій' @('Час','Користувач','Дія','ID / Аркуш') 'D' 1005
  $audit.Columns.Item('A').ColumnWidth=25
  $audit.Columns.Item('B').ColumnWidth=24
  $audit.Columns.Item('C').ColumnWidth=24
  $audit.Columns.Item('D').ColumnWidth=35
  $bill=$book.Worksheets.Item('Нарахування')
  $billHeaders=@('Статус','Дата закриття','Користувач','Маршрут ID','Авто ID','Прогноз PLN','Факт км','Прогноз л','Ціна PLN/л','План км')
  for($i=0;$i -lt $billHeaders.Count;$i++){$bill.Cells.Item(12,$i+12).Value2=$billHeaders[$i]}
  $bill.Range('L12:U12').Font.Bold=$true
  $bill.Range('L12:U12').Font.Name='Times New Roman'
  $bill.Range('L12:U12').Font.Size=14
  Set-NumberFormatSafe ($bill.Range('M13:M132')) 'dd.mm.yyyy hh:mm' 'M13:M132'
  $bill.Range('O13:U132').Interior.Color=14277081
  # v10 application DB. Append-only records: no per-driver and per-trip template limit.
  # No web calls, no scheduled refreshes, everything is entered by the administrator.
  Write-Host '3/5. Динамічні реєстри, 150 паркінгів і центральна панель адміністратора...'
  $trips=Catalog $book 'TC_Рейси' @('ID','Створено','Водій ID','Авто ID','Маршрут ID','Клієнт','Виїзд план','Статус','Виїзд факт','Прибуття факт','Одометр поч','Одометр кін','Км факт','Дохід USD','Отримано USD','Км план','Прогноз літрів','Прогноз PLN','Витрати USD','Прибуток USD','Примітка','Час завершення','Кермування год','Простій год','Втрати год','Прибуття план','Одометр зараз','Дата оновлення') 'AB' 55
  $trips.Range('A6:Z55').Interior.Color=14277081
  $trips.Columns.Item('A').ColumnWidth=31
  $trips.Columns.Item('C').ColumnWidth=19
  $trips.Columns.Item('D').ColumnWidth=17
  $trips.Columns.Item('E').ColumnWidth=19
  $trips.Columns.Item('F').ColumnWidth=27
  $trips.Columns.Item('H').ColumnWidth=19
  Set-NumberFormatSafe ($trips.Range('B6:B55')) 'dd.mm.yyyy hh:mm' 'TC_Рейси creation dates'
  Set-NumberFormatSafe ($trips.Range('G6:J55')) 'dd.mm.yyyy' 'TC_Рейси trip dates'
  $spend=Catalog $book 'TC_Витрати' @('Створено','Рейс ID','Категорія','Паркінг','Сума','Валюта','USD/PLN','Сума USD','Дата витрати','Коментар','Користувач') 'K' 55
  $spend.Range('A6:K55').Interior.Color=14277081
  $spend.Columns.Item('B').ColumnWidth=29
  $spend.Columns.Item('C').ColumnWidth=20
  $spend.Columns.Item('J').ColumnWidth=44
  $park=Catalog $book 'TC_Паркінги' @('Код','Назва пакета','Локація','Тривалість','Тариф PLN/пакет','Джерело/коментар','Статус') 'G' 155
  $park.Range('A6:G155').Interior.Color=14348287
  $park.Columns.Item('B').ColumnWidth=38
  $park.Columns.Item('F').ColumnWidth=60
  $park.Range('A3').Value2='150 підготовлених місць. Новий паркінг створюється у вікні адміністратора; при заповненні діапазон розширюється.'
  # Copy locally available parking tariff packages from the historical source
  # into the administrator catalog; no Internet request, no automatic tariff update.
  $legacyParks=$book.Worksheets.Item('Паркінги')
  $parkCount=0
  for($srcRow=11;$srcRow -le 160;$srcRow++){
    $parkCode=[string]$legacyParks.Cells.Item($srcRow,1).Value2
    if([string]::IsNullOrWhiteSpace($parkCode)){continue}
    $pr=6+$parkCount
    $park.Cells.Item($pr,1).Value2=$parkCode
    $park.Cells.Item($pr,2).Value2=[string]$legacyParks.Cells.Item($srcRow,2).Value2
    $park.Cells.Item($pr,3).Value2=[string]$legacyParks.Cells.Item($srcRow,2).Value2
    $park.Cells.Item($pr,4).Value2=[string]$legacyParks.Cells.Item($srcRow,14).Value2
    $amount=$legacyParks.Cells.Item($srcRow,3).Value2
    if($null -ne $amount -and [string]$amount -ne ''){Set-ExcelNumber ($park.Cells.Item($pr,5)) ([double]$amount) ('Parking '+$parkCode)}
    $park.Cells.Item($pr,6).Value2=[string]$legacyParks.Cells.Item($srcRow,6).Value2
    $park.Cells.Item($pr,7).Value2='Попередній тариф / перевірити'
    $parkCount++
    if($parkCount -ge 150){break}
  }
  Write-Host ('Паркінги: '+$parkCount+' історичних пакетів внесено, до 150 місць попередньо підготовлено.')
  $activity=Catalog $book 'TC_Події' @('Час','Користувач','Операція',"Об'єкт") 'D' 55
  # Read-only operational screen: not part of the CRUD sheets.
  $live=$book.Worksheets.Add()
  $live.Name='Поточний стан'
  $live.Range('A1').Value2='ТРАНСПОРТНИЙ ЦЕНТР | ПОТОЧНИЙ СТАН'
  $live.Range('A3').Value2='Дані оновлюються автоматично при відкритті книги та переході на аркуш; можна оновити кнопкою.'
  $live.Range('A5').Value2='АКТИВНІ РЕЙСИ'
  $live.Range('A7').Value2='Рейс ID'
  $live.Range('B7').Value2='Водій'
  $live.Range('F7').Value2='Днів'
  $live.Range('I7').Value2='Витрати USD'
  $live.Columns.Item('A').ColumnWidth=28
  $live.Columns.Item('B').ColumnWidth=25
  $live.Range('A1:K35').Font.Name='Times New Roman'
  $live.Range('A1:K35').Font.Size=12
  $live.Range('A1').Font.Size=18
  $live.Range('A1').Font.Bold=$true
  DashboardButton $live 'На головну' 'TC_GoHome' 'J4'
  $liveButton=$live.Shapes.Item($live.Shapes.Count)
  $liveButton.OnAction="'Транспортний_центр.xlsm'!TransportInsights.TC_GoHome"
  DashboardButton $live 'Оновити стан' 'TC_ShowLive' 'F4'
  $live.Shapes.Item($live.Shapes.Count).OnAction="'Транспортний_центр.xlsm'!TransportStatusFX.TC_ShowLive"

  # Archive and FX are independent visible screens, not hidden ledger worksheets.
  $archive=$book.Worksheets.Add()
  $archive.Name='Архів рейсів'
  $archive.Range('A1').Value2='ТРАНСПОРТНИЙ ЦЕНТР | АРХІВ РЕЙСІВ'
  $archive.Range('A3').Value2='Дані архіву оновлюються автоматично при відкритті та переході на аркуш.'
  $archive.Range('A5').Value2='ЗАВЕРШЕНІ РЕЙСИ'
  $archive.Range('A7').Value2='Рейс ID'
  $archive.Range('B7').Value2='Водій'
  $archive.Range('G7').Value2='Днів'
  $archive.Range('I7').Value2='Витрати USD'
  $archive.Range('A1:L35').Font.Name='Times New Roman'
  $archive.Columns.Item('A').ColumnWidth=25
  DashboardButton $archive 'На головну' 'TC_GoHome' 'J4'
  $archive.Shapes.Item($archive.Shapes.Count).OnAction="'Транспортний_центр.xlsm'!TransportInsights.TC_GoHome"
  DashboardButton $archive 'Оновити архів' 'TC_ShowArchive' 'F4'
  $archive.Shapes.Item($archive.Shapes.Count).OnAction="'Транспортний_центр.xlsm'!TransportStatusFX.TC_ShowArchive"
  $rates=$book.Worksheets.Add()
  $rates.Name='Курси валют'
  $rates.Range('A1:F2').Merge()
  $rates.Range('A1').Value2='ТРАНСПОРТНИЙ ЦЕНТР | КУРСИ ВАЛЮТ'
  $rates.Range('A1:F2').Interior.Color=(13+36*256+31*65536)
  $rates.Range('A1:F2').Font.Color=(245+233*256+196*65536)
  $rates.Range('A1:F2').Font.Size=20
  $rates.Range('A1:F2').Font.Bold=$true
  $rates.Range('A1:F2').RowHeight=25
  $rates.Range('A3:F3').Merge()
  $rates.Range('A3').Value2='Курс: скільки PLN коштує 1 одиниця валюти. Порожній курс = конвертація недоступна. Оновлення тільки вручну.'
  $rates.Range('A3').Font.Italic=$true
  $rates.Range('A1:F28').Font.Name='Times New Roman'
  $rates.Range('A1:F28').Font.Size=13
  $rates.Columns.Item('A').ColumnWidth=17
  $rates.Columns.Item('B').ColumnWidth=27
  $rates.Columns.Item('C').ColumnWidth=22
  $rates.Columns.Item('D').ColumnWidth=29
  $rates.Columns.Item('E').ColumnWidth=24
  $rates.Columns.Item('F').ColumnWidth=20
  $rateHeaders=@('Валюта','PLN за 1 одиницю','Дата курсу','Джерело','Оновлено')
  for($i=0;$i -lt $rateHeaders.Count;$i++){$rates.Cells.Item(5,$i+1).Value2=$rateHeaders[$i]}
  $rates.Range('A5:E5').Interior.Color=(39+91*256+75*65536)
  $rates.Range('A5:E5').Font.Color=16777215
  $rates.Range('A5:E5').Font.Bold=$true
  $rateCodes=@('PLN','USD','EUR','UAH')
  for($i=0;$i -lt $rateCodes.Count;$i++){$rates.Cells.Item(6+$i,1).Value2=$rateCodes[$i]}
  $rates.Range('B6').Formula='=1'
  $rates.Range('D6').Value2='Базова валюта'
  $rates.Range('B7:B9').ClearContents() # never assert stale market rates
  $rates.Range('B7:B9').Interior.Color=(255+242*256+174*65536) # editable yellow
  Set-NumberFormatSafe ($rates.Range('B6:B9')) '0.000000' 'Currency quote precision'
  Set-NumberFormatSafe ($rates.Range('C7:C9')) 'dd.mm.yyyy' 'Currency date'
  $rates.Range('A12:F12').Merge()
  $rates.Range('A12').Value2='Введіть USD/EUR/UAH у жовті поля, потім натисніть Застосувати ручні курси.'
  $rates.Range('A12').Font.Bold=$true
  $rates.Range('A14:F15').Merge()
  $rates.Range('A14').Value2='Це курси для перегляду еквівалентів. Облік витрат фіксується в USD під час внесення квитанції. Немає автоматичних запитів до Інтернету.'
  $rates.Range('A14').WrapText=$true
  DashboardButton $rates 'Оновити NBP' 'TC_RatesFetchNBP' 'A18'
  $rates.Shapes.Item($rates.Shapes.Count).OnAction="'Транспортний_центр.xlsm'!TransportStatusFX.TC_RatesFetchNBP"
  DashboardButton $rates 'Застосувати ручні курси' 'TC_RatesManualApply' 'D18'
  $rates.Shapes.Item($rates.Shapes.Count).OnAction="'Транспортний_центр.xlsm'!TransportStatusFX.TC_RatesManualApply"
  DashboardButton $rates 'На головну' 'TC_GoHome' 'A22'
  $rates.Shapes.Item($rates.Shapes.Count).OnAction="'Транспортний_центр.xlsm'!TransportInsights.TC_GoHome"
  # All online actions remain exclusively button-initiated.

  $activity.Columns.Item('C').ColumnWidth=30
  $activity.Columns.Item('D').ColumnWidth=42
  $activity.Range('A6:D55').Interior.Color=14277081
  # Route KPIs are now bound to the expandable operational ledger, not the legacy 120-trip sheet.
  $rt.Range('G6').Formula=(ExcelExpr '=IF(A6="","",COUNTIF(~TC_Рейси~!$E$6:$E$50000,A6))')
  $rt.Range('H6').Formula=(ExcelExpr '=IF(A6="","",SUMIF(~TC_Рейси~!$E$6:$E$50000,A6,~TC_Рейси~!$M$6:$M$50000))')
  $rt.Range('I6').Formula=(ExcelExpr '=IF(A6="","",COUNTIFS(~TC_Рейси~!$E$6:$E$50000,A6,~TC_Рейси~!$H$6:$H$50000,"Завершено"))')
  foreach($letter in @('G','H','I')){$rt.Range("$($letter)6:$($letter)205").FillDown()}
  $index=$book.Worksheets.Item('Зміст')
  # v8 has merged cells A1:H2, A3:H3, A5:H5 etc. Remove BEFORE merging dashboard.
  $index.Range('A1:R80').UnMerge()
  # Replace the legacy tab-oriented content page with a compact forest-themed dashboard.
  $index.Range('A1:R80').ClearContents()
  # RGB(13,32,28) -- quiet dark forest; previous palette was inadvertently orange.
  $index.Range('A1:R80').Interior.Color=(13+32*256+28*65536)
  $index.Range('A1:R80').Font.Name='Times New Roman'
  $index.Range('A1:R80').Font.Size=13
  $index.Range('A1:R80').RowHeight=17
  $index.Rows.Item(2).RowHeight=48
  $index.Range('A1:R80').Font.Color=(225+237*256+230*65536)
  foreach($col in @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R')){$index.Columns.Item($col).ColumnWidth=11.5}
  $index.Rows.Item(1).RowHeight=30
  $index.Rows.Item(2).RowHeight=48
  $index.Range('B2:Q3').Merge()
  $index.Range('B2').Value2='ТРАНСПОРТНИЙ ЦЕНТР | АДМІНІСТРАТОР'
  $index.Range('B2').Font.Bold=$true
  $index.Range('B2').Font.Size=22
  $index.Range('B2').Font.Color=(245+233*256+196*65536)
  $index.Range('B5:Q5').Merge()
  $index.Range('B5').Value2='Створення й ведення рейсів, автомобілі, контакти, маршрути та витрати — лише через вікна.'
  $index.Range('B5').Font.Color=(170+199*256+184*65536)
  # KPI text must not spill/clamp into a 17pt row with 23pt glyphs.
  # Merge dedicated two-row value cards without shifting dashboard buttons.
  # Excel layout remains readable at 85% zoom and on smaller displays.
  $index.Rows.Item(7).RowHeight=21
  $index.Rows.Item(8).RowHeight=20
  $index.Rows.Item(11).RowHeight=21
  $index.Rows.Item(12).RowHeight=20
  foreach($group in @('B7:D7','F7:H7','J7:L7','N7:P7','B11:D11','F11:H11','J11:L11','N11:P11')){
    $index.Range($group).Merge()
    $index.Range($group).HorizontalAlignment=-4108 # xlCenter
    $index.Range($group).VerticalAlignment=-4108
  }
  foreach($group in @('B8:D9','F8:H9','J8:L9','N8:P9','B12:D13','F12:H13','J12:L13','N12:P13')){
    $index.Range($group).Merge()
    $index.Range($group).HorizontalAlignment=-4108
    $index.Range($group).VerticalAlignment=-4108
  }
  $kpi=@(
    @('B7','ВОДІЇВ','B8','=COUNTA(Контакти_водіїв!$A$6:$A$1000)'),
    @('F7','ЗАПЛАНОВАНО','F8','=COUNTIF(~TC_Рейси~!$H$6:$H$50000,"Заплановано")'),
    @('J7','У РЕЙСІ','J8','=COUNTIF(~TC_Рейси~!$H$6:$H$50000,"У рейсі")'),
    @('N7','ЗАВЕРШЕНО','N8','=COUNTIF(~TC_Рейси~!$H$6:$H$50000,"Завершено")'),
    @('B11','МАРШРУТІВ','B12','=COUNTA(Маршрути_Європа!$A$6:$A$1000)'),
    @('F11','ПАРКІНГІВ','F12','=COUNTA(~TC_Паркінги~!$A$6:$A$1000)'),
    @('J11','ВИТРАТИ USD','J12','=SUM(~TC_Витрати~!$H$6:$H$50000)'),
    @('N11','ДОХІД USD','N12','=SUM(~TC_Рейси~!$N$6:$N$50000)')
  )
  foreach($def in $kpi){
    $label=[string]$def[0];$value=[string]$def[2]
    $index.Range($label).Value2=[string]$def[1]
    $index.Range($label).Font.Bold=$true
    $index.Range($label).Font.Color=11259375
    $index.Range($value).Formula=(ExcelExpr ([string]$def[3]))
    $index.Range($value).Font.Size=21
    $index.Range($value).Font.Bold=$true
    $index.Range($value).Font.Color=16777215
  }
  $index.Range('B15:Q15').Merge()
  $index.Range('B15').Value2='РОБОТА З РЕЙСАМИ'
  $index.Range('B15').Font.Bold=$true
  $index.Range('B15').Font.Color=(245+233*256+196*65536)
  $index.Range('B20:Q20').Merge()
  $index.Range('B20').Value2='ДОВІДНИКИ ТА КОНТРОЛЬ'
  $index.Range('B20').Font.Bold=$true
  $index.Range('B20').Font.Color=(245+233*256+196*65536)
  $index.Range('B25:Q25').Merge()
  $index.Range('B25').Value2='ФІНАНСИ ТА СИСТЕМА'
  $index.Range('B25').Font.Bold=$true
  $index.Range('B25').Font.Color=(245+233*256+196*65536)
  $index.Range('B36:Q36').Merge()
  $index.Range('B36').Value2='Робочі журнали приховані. Поточний стан і сформовані звіти доступні для перегляду.'
  $index.Range('B36').Font.Color=(170+199*256+184*65536)
  # Remove legacy shapes from the old navigation dashboard before placing new buttons.
  for($shapeIndex=$index.Shapes.Count;$shapeIndex -ge 1;$shapeIndex--){
    try{$index.Shapes.Item($shapeIndex).Delete()}catch{throw ('Old dashboard shape cannot be removed: '+$_.Exception.Message)}
  }
  # Deselect frozen panes in EVERY worksheet. FreezePanes is a Window property.
  foreach($sheet in $book.Worksheets){
    $sheet.Activate()
    try { $excel.ActiveWindow.FreezePanes=$false } catch { Write-Warning ('No freeze access: '+$sheet.Name) }
    try { $excel.ActiveWindow.SplitRow=0; $excel.ActiveWindow.SplitColumn=0 } catch { }
  }
  $index.Activate()

  # CLEAN: no fixture loading, all transactional records remain blank.
  Write-Host 'CLEAN BUILD: no test records inserted.' -ForegroundColor Green
  Write-Host '4/5. Вбудовування VBA та кнопок...'
  $sourceVBA=[System.IO.File]::ReadAllText($moduleFile,[System.Text.Encoding]::UTF8)
  $sourceVBA=$sourceVBA -replace '(?m)^Attribute VB_Name = "[^"]*"\r?\n',''
  try {
     $targetProject=$book.VBProject
     if($null -eq $targetProject){throw 'VBProject основної книги недоступний.'}
     $targetComponents=$targetProject.VBComponents
     if($null -eq $targetComponents){throw 'VBComponents основної книги недоступний.'}
     $component=$targetComponents.Add(1)
     if($null -eq $component){throw 'Excel не повернув об`єкт щойно створеного модуля.'}
     $component.Name='TransportCenter'
     $codeModule=$component.CodeModule
     if($null -eq $codeModule){throw 'Excel не повернув CodeModule основної книги.'}
     $codeModule.AddFromString($sourceVBA)
     $appSource=[System.IO.File]::ReadAllText($appModuleFile,[System.Text.Encoding]::UTF8)
     $appSource=$appSource -replace '(?m)^Attribute VB_Name = "[^"]*"\r?\n',''
     $appComponent=$targetComponents.Add(1)
     Check ($null -ne $appComponent) 'Немає VBA-модуля TransportApp.'
     $appComponent.Name='TransportApp'
     $appComponent.CodeModule.AddFromString($appSource)
     # Independent reporting engine. Keep the proven v9 save/close mechanism intact.
     $insightsSource=[System.IO.File]::ReadAllText($insightsFile,[System.Text.Encoding]::UTF8)
     $insightsSource=$insightsSource -replace '(?m)^Attribute VB_Name = "[^"]*"\r?\n',''
     $insightsComponent=$targetComponents.Add(1)
     Check ($null -ne $insightsComponent) 'Немає VBA-модуля аналітики.'
     $insightsComponent.Name='TransportInsights'
     $insightsComponent.CodeModule.AddFromString($insightsSource)
     $statusFxSource=[System.IO.File]::ReadAllText($statusFxFile,[System.Text.Encoding]::UTF8)
     $statusFxSource=$statusFxSource -replace '(?m)^Attribute VB_Name = "[^"]*"\r?\n',''
     $statusFxComponent=$targetComponents.Add(1)
     Check ($null -ne $statusFxComponent) 'Немає VBA-модуля статусів та курсів.'
     $statusFxComponent.Name='TransportStatusFX'
     $statusFxComponent.CodeModule.AddFromString($statusFxSource)
     # Local-only refresh on opening or activating either operational view.
     # Event code runs inside the already-embedded VBA project (no VBProject trust at runtime).
     try {
       $workbookComponent=$null
       $workbookCodeName=[string]$book.CodeName
       if(-not [string]::IsNullOrWhiteSpace($workbookCodeName)) {
         try { $workbookComponent=$targetComponents.Item($workbookCodeName) } catch { }
       }
       if($null -eq $workbookComponent){$workbookComponent=$targetComponents.Item('ThisWorkbook')}
       Check ($null -ne $workbookComponent) 'Не знайдено ThisWorkbook для автоматичного оновлення.'
       $eventSource=[System.IO.File]::ReadAllText((Join-Path $dir 'vba\TransportWorkbookEvents.txt'),[System.Text.Encoding]::UTF8)
       $workbookComponent.CodeModule.AddFromString($eventSource)
       Write-Host 'Автоматичний офлайн-перерахунок: Workbook_Open / SheetActivate додано.' -ForegroundColor Green
     } catch {
       # Fail open: VBA UI and manual Refresh buttons still work even when
       # Office policy rejects the optional workbook event handler.
       Write-Warning ('Автоматичний перерахунок не вдалося додати: '+$_.Exception.Message+'. Кнопки оновлення залишаються доступними.')
     }
     # Microsoft Forms 2.0 design created ONCE during build; no runtime VBProject permissions.
     $vbaStage='UserForm: VBComponents.Add(3)'
     $uiComponent=$targetComponents.Add(3) # vbext_ct_MSForm
     Check ($null -ne $uiComponent) 'Excel не може створити UserForm.'
     $uiComponent.Name='frmTCAdmin'
     $vbaStage='UserForm: Designer and design-time properties'
     $designer=$uiComponent.Designer
     Check ($null -ne $designer) 'Компонент UserForm не надав Designer.'
     # VBIDE.UserForm Designer is NOT guaranteed to expose Width/Height via
     # PowerShell COM. Set design-time size and caption via VBComponent.Properties.
     # This is the documented interface for programmatically-created UserForms.
     $uiComponent.Properties.Item('Caption').Value='Транспортний центр | Адміністратор'
     $uiComponent.Properties.Item('Width').Value=710
     $uiComponent.Properties.Item('Height').Value=575
     function FormCtrl([string]$kind,[string]$name,[string]$caption,[int]$left,[int]$top,[int]$width,[int]$height){
       $ctl=$designer.Controls.Add(('Forms.'+$kind+'.1'),$name,$true)
       if($null -eq $ctl){throw ('Не створено поле форми: '+$name)}
       # Microsoft Forms Controls support Move(left, top, width, height).
       # Include the control identifier in diagnostics if Office rejects placement.
       try { $ctl.Move([single]$left,[single]$top,[single]$width,[single]$height) }
       catch { throw ('Помилка геометрії поля '+$name+' ('+$kind+'): '+$_.Exception.Message) }
       try{$ctl.Font.Name='Times New Roman';$ctl.Font.Size=12}catch{}
       if($kind -eq 'Label' -or $kind -eq 'CommandButton'){$ctl.Caption=$caption}
       if($kind -eq 'ComboBox'){$ctl.Style=2} # fmStyleDropDownList, list-only selection
       return $ctl
     }
     $vbaStage='UserForm: create MSForms controls'
     FormCtrl 'Label' 'lblTitle' 'ТРАНСПОРТНИЙ ЦЕНТР' 25 15 650 28|Out-Null
     FormCtrl 'Label' 'lblInfo' 'Адміністратор веде всі операції.' 25 45 650 22|Out-Null
     $comboItems=@('Driver','Vehicle','Route','Trip','Parking','Kind','Currency')
     $comboCaptions=@('Водій','Автомобіль','Маршрут','Рейс','Паркінг','Категорія','Валюта')
     for($i=0;$i -lt $comboItems.Count;$i++){
       $top=86+$i*49
       FormCtrl 'Label' ('lbl'+$comboItems[$i]) $comboCaptions[$i] 25 $top 270 18|Out-Null
       FormCtrl 'ComboBox' ('cbo'+$comboItems[$i]) '' 25 ($top+20) 283 25|Out-Null
     }
     for($i=1;$i -le 8;$i++){
       $top=86+($i-1)*43
       FormCtrl 'Label' ('lbl'+$i) ('Поле '+$i) 335 $top 325 17|Out-Null
       FormCtrl 'TextBox' ('txt'+$i) '' 335 ($top+18) 325 22|Out-Null
     }
     FormCtrl 'Label' 'lblPreview' 'Перед змінами збережеться резервна копія.' 25 456 635 25|Out-Null
     FormCtrl 'CommandButton' 'cmdSave' 'Зберегти' 335 498 159 33|Out-Null
     FormCtrl 'CommandButton' 'cmdCancel' 'Скасувати' 505 498 155 33|Out-Null
     $formHandlers=@'
Option Explicit
Private mInitialized As Boolean
Private Sub UserForm_Activate()
    ' An instance created through VBIDE must finish activating before lists are loaded.
    ' Prevent repeated AddItem calls if Excel reactivates the modal form.
    If mInitialized Then Exit Sub
    mInitialized = True
    TransportApp.TC_FormInit Me
End Sub
Private Sub cmdSave_Click()
    TransportApp.TC_FormSave Me
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cboDriver_Change()
    TransportApp.TC_FormSelectionChanged Me, "Driver"
End Sub
Private Sub cboVehicle_Change()
    TransportApp.TC_FormSelectionChanged Me, "Vehicle"
End Sub
Private Sub cboRoute_Change()
    TransportApp.TC_FormSelectionChanged Me, "Route"
End Sub
Private Sub cboParking_Change()
    TransportApp.TC_FormSelectionChanged Me, "Parking"
End Sub
Private Sub cboKind_Change()
    TransportApp.TC_FormSelectionChanged Me, "Kind"
End Sub
'@
     $vbaStage='UserForm: insert event handlers'
     $uiComponent.CodeModule.AddFromString($formHandlers)
  } catch {
     throw ('VBA_IMPORT_FAILED [' + [string]$vbaStage + ']: початкову перевірку дозволу пройдено, але під час внесення VBA у робочу книгу виникла помилка: ' + $_.Exception.Message)
  }
  # Buttons launch MODAL USERFORMS; NEVER activate a data worksheet.
  # Preserve the main tab as the ONLY visible operational surface.
  $index=$book.Worksheets.Item('Зміст')
  DashboardButton $index 'Створити рейс' 'TC_UI_NewTrip' 'B17'
  DashboardButton $index 'Розпочати рейс' 'TC_UI_StartTrip' 'F17'
  DashboardButton $index 'Завершити рейс' 'TC_UI_FinishTrip' 'J17'
  DashboardButton $index 'Внести витрати' 'TC_UI_Expense' 'N17'
  DashboardButton $index 'Автопарк' 'TC_UI_Fleet' 'B22'
  DashboardButton $index 'Водії / контакти' 'TC_UI_Driver' 'F22'
  DashboardButton $index 'Маршрути Європи' 'TC_UI_Route' 'J22'
  DashboardButton $index 'Паркінги' 'TC_UI_Parking' 'N22'
  DashboardButton $index 'Фінансовий звіт' 'TC_UI_Report' 'B27'
  DashboardButton $index 'Резервна копія' 'TC_UI_Backup' 'F27'
  DashboardButton $index 'Оновити NBP вручну' 'TC_UI_OnlineNBP' 'J27'
  DashboardButton $index 'Поточний стан' 'TC_ShowLive' 'N27'
  DashboardButton $index 'Оновити пробіг' 'TC_UI_Progress' 'B32'
  DashboardButton $index 'Архів рейсів' 'TC_ShowArchive' 'J32'
  DashboardButton $index 'Курси валют' 'TC_ShowRates' 'N32'
  DashboardButton $index 'Відновити копію' 'TC_SeeBackup' 'F32'
  # Assign new dashboard buttons to TransportApp; legacy restore uses TransportCenter.
  $buttonNames=@('Створити рейс','Розпочати рейс','Завершити рейс','Внести витрати','Автопарк','Водії / контакти','Маршрути Європи','Паркінги','Фінансовий звіт','Резервна копія','Оновити NBP вручну','Оновити пробіг')
  foreach($sh in $index.Shapes){
    $caption=''
    try {$caption=[string]$sh.TextFrame.Characters().Text} catch {continue}
    if($buttonNames -contains $caption){
      $oldAction=[string]$sh.OnAction
      $sh.OnAction=$oldAction.Replace('TransportCenter.','TransportApp.')
    }
    $sh.Placement=3 # xlFreeFloating; do not move/resize with cells
  }
  foreach($shape in $index.Shapes){
    $caption=''
    try { $caption=[string]$shape.TextFrame.Characters().Text } catch { continue }
    if($caption -eq 'Поточний стан'){$shape.OnAction="'Транспортний_центр.xlsm'!TransportStatusFX.TC_ShowLive"}
    if($caption -eq 'Архів рейсів'){$shape.OnAction="'Транспортний_центр.xlsm'!TransportStatusFX.TC_ShowArchive"}
    if($caption -eq 'Курси валют'){$shape.OnAction="'Транспортний_центр.xlsm'!TransportStatusFX.TC_ShowRates"}
    if($caption -eq 'Оновити NBP вручну'){$shape.OnAction="'Транспортний_центр.xlsm'!TransportStatusFX.TC_RatesFetchNBP"}
  }
  if($book.Worksheets.Item(1).Name -ne 'Зміст'){$index.Move($book.Worksheets.Item(1))}
  # Other worksheets stay in the file for formulas/audit, but are hidden from customers.
  $index.Activate()
  foreach($other in $book.Worksheets){
    if($other.Name -ne 'Зміст' -and $other.Name -ne 'Поточний стан' -and $other.Name -ne 'Архів рейсів' -and $other.Name -ne 'Курси валют'){$other.Visible=0}
  }
  try {$excel.ActiveWindow.DisplayGridlines=$false; $excel.ActiveWindow.Zoom=85; $excel.ActiveWindow.ScrollRow=1; $excel.ActiveWindow.ScrollColumn=1} catch {}
  foreach($view in @($index,$live,$archive,$rates)) {
    try {
      $view.Activate()
      $excel.ActiveWindow.DisplayGridlines=$false
      $excel.ActiveWindow.FreezePanes=$false
      $excel.ActiveWindow.Zoom=85
      $excel.ActiveWindow.ScrollRow=1
      $excel.ActiveWindow.ScrollColumn=1
    } catch { Write-Warning ('Could not configure view for '+$view.Name+': '+$_.Exception.Message) }
  }
  $index.Activate()
  Write-Host ('Інтерфейс: панель адміністратора, поточний стан, архів і курси валют. '+($book.Worksheets.Count-4)+' технічних вкладок приховано.')
  Write-Host '5/5. Збереження повного XLSM...'
  try {
    $excel.Calculation=-4105
  } catch {
    Write-Warning 'Excel did not accept automatic calculation mode; trying a single calculation before save.'
    try { $excel.Calculate() } catch { Write-Warning 'Manual recalculation is available with F9.' }
  }
  # Refuse incorrectly seeded releases. Reference routes are NOT transactions.
  foreach($entry in @(
    @('TC_Рейси','A6:A50000'), @('TC_Витрати','B6:B50000'),
    @('Автопарк','A6:A1000'), @('Контакти_водіїв','A6:A1000'),
    @('Реєстр_водіїв','A13:A52'), @('TC_Паркінги','A6:A1000'),
    @('Журнал_операцій','A6:A1000')
  )){
    $ws=$book.Worksheets.Item([string]$entry[0])
    Check ([double]$excel.WorksheetFunction.CountA($ws.Range([string]$entry[1])) -eq 0) ('Clean build has unwanted records: ' + $entry[0])
  }
  Write-Host 'CLEAN DATA CHECK: 0 drivers, vehicles, trips, receipts, parkings, or operations.' -ForegroundColor Green
  # Save the actual macro-enabled workbook BEFORE running any injected VBA.
  # Previous installers ran code while the workbook was still an .xlsx container.
  $index.Activate()
  $book.SaveAs($target,52)
  Check (Test-Path -LiteralPath $target) 'Excel не створив нову книгу XLSM.'
  # Strict VBA runtime smoke test, after the complete XLSM is present.
  $refreshReady = $false
  $refreshFailure = ''
  try {
    $runName = "'" + $book.Name + "'!TransportStatusFX.TC_RefreshBoth"
    [void]$excel.Run($runName)
    $refreshReady = $true
    Write-Host 'CLEAN status/archive smoke test AFTER SaveAs XLSM: SUCCESS' -ForegroundColor Green
  } catch {
    $refreshFailure = $_.Exception.ToString()
    Write-Warning ('VBA smoke test FAILED after SaveAs. XLSM is only for diagnostics: ' + $_.Exception.Message)
  }
  # Preserve both the recalculated data and/or diagnostic cell N1 for inspection.
  $index.Activate()
  $book.Save()
  # SaveAs does NOT close the workbook. File.OpenRead while Excel still owns the
  # XLSM may throw: 'The process cannot access the file ... another process'.
  # Close ONLY the workbook opened by our private Excel COM instance, then quit.
  Write-Host '5/5. Книгу XLSM збережено. Закриваємо службовий Excel перед перевіркою файлу...'
  $book.Close($false)
  $book=$null
  if ($guard -ne $null) { $guard.Close($false); $guard=$null }
  $excel.Quit()
  $excel=$null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  [GC]::Collect()

  Add-Type -AssemblyName System.IO.Compression
  # Windows Defender / file indexing can momentarily open a newly created file.
  # Use shared read and bounded retries; do not trigger another Office instance.
  $fs=$null
  $lastReadError=$null
  for($attempt=1;$attempt -le 12;$attempt++) {
    try {
      $fs=[System.IO.File]::Open($target,[System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,[System.IO.FileShare]::ReadWrite)
      break
    } catch [System.IO.IOException] {
      $lastReadError=$_.Exception.Message
      Start-Sleep -Milliseconds 450
    }
  }
  if($null -eq $fs){
    throw ('Книгу XLSM створено і Excel закрито, але ОС поки не дозволяє перевірити файл: ' + $lastReadError + '. Готову книгу НЕ видалено: ' + $target)
  }
  try {
    $arc=[System.IO.Compression.ZipArchive]::new($fs,[System.IO.Compression.ZipArchiveMode]::Read)
    try {
      Check ($null -ne $arc.GetEntry('xl/vbaProject.bin')) 'Книгу XLSM створено, але архів не містить VBA-проєкт.'
      Check ($null -ne $arc.GetEntry('xl/workbook.xml')) 'В XLSM немає структури Excel.'
      Check ($null -ne $arc.GetEntry('[Content_Types].xml')) 'В XLSM немає опису вмісту Open XML.'
      Write-Host 'Перевірка файлу XLSM успішна: VBA-проєкт та структура книги присутні.' -ForegroundColor Green
    } finally {$arc.Dispose()}
  } finally {$fs.Dispose()}
  $backup=Join-Path $dir '_Резервні_копії'
  if(-not(Test-Path -LiteralPath $backup)){[void](New-Item -Path $backup -ItemType Directory)}
  if (-not $refreshReady) {
      $diagnostic = Join-Path $dir 'VBA_SMOKE_TEST_FAILED.txt'
      @('Transport Center CLEAN v10.14.7: VBA refresh smoke test FAILED after SaveAs.',
        ('Time: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
        ('Workbook: '+$target),
        ('Details: '+$refreshFailure),
        'The XLSM exists for debugging ONLY, not a verified working install.',
        'Inspect Поточний стан!N1 or Архів рейсів!N1 for the original failure stage.') |
          Out-File -LiteralPath $diagnostic -Encoding UTF8
      throw ('VBA_REFRESH_FAILED: status/archive engine smoke test failed. ' +
             'See VBA_SMOKE_TEST_FAILED.txt; do not use XLSM for accounting.')
  }
  Write-Host 'ГОТОВО: Транспортний_центр.xlsm — програма адміністратора v10' -ForegroundColor Green
  Write-Host 'Джерело коду лише локальні файли. GitHub не використовується.'
  Write-Host 'Примітка: структуру XLSM перевірено; виконання кнопок VBA ще потрібно перевірити в Excel LTSC.'
} catch {
  $failure = $_
  $errorFile = Join-Path $dir 'INSTALL_ERROR.txt'
  $diagnostics = @(
    'Transport Center Offline v10 - Excel installation diagnostics',
    ('Date: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')),
    ('Error: ' + $failure.Exception.Message),
    ('Line: ' + $failure.InvocationInfo.ScriptLineNumber),
    ('Code: ' + [string]$failure.InvocationInfo.Line),
    ('Position: ' + [string]$failure.InvocationInfo.PositionMessage),
    ('HResult: ' + $failure.Exception.HResult),
    ('ErrorType: ' + $failure.Exception.GetType().FullName),
    ('Script: ' + $failure.InvocationInfo.ScriptName)
  )
  $diagnostics | Out-File -LiteralPath $errorFile -Encoding UTF8
  Write-Host ('ПОМИЛКА: ' + $failure.Exception.Message) -ForegroundColor Red
  Write-Host ('Рядок: ' + $failure.InvocationInfo.ScriptLineNumber) -ForegroundColor Yellow
  Write-Host ('Деталі збережено: ' + $errorFile) -ForegroundColor Yellow
  Write-Host 'Вихідна XLSX не перезаписувалась.' -ForegroundColor Yellow
  exit 1
} finally {
  if($book -ne $null){try{$book.Close($false)}catch{}}
  if($guard -ne $null){try{$guard.Close($false)}catch{}}
  if($excel -ne $null){try{$excel.Quit()}catch{}}
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}