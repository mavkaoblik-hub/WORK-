param(
 [string]$Source = "Транспортний_центр_v8_робочий_Times_New_Roman.xlsx",
 [string]$Destination = "Транспортний_центр.xlsm"
)
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourcePath = Join-Path $dir $Source
$destinationPath = Join-Path $dir $Destination
$modulePath = Join-Path $dir 'vba\TransportCenter.bas'
if (!(Test-Path -LiteralPath $sourcePath)) { throw "Не знайдено вихідний XLSX: $sourcePath" }
if (!(Test-Path -LiteralPath $modulePath)) { throw "Не знайдено VBA-модуль: $modulePath" }
if (Test-Path -LiteralPath $destinationPath) { throw "Файл $Destination вже існує. Перейменуйте його або збережіть окремо. Перезапис заборонено." }
$excel = $null; $book = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.Calculation = -4135 # xlCalculationManual: prevent web formulas evaluating on source open
  $excel.AutomationSecurity = 3 # disable macros while importing source
  $excel.EnableEvents = $false
  $excel.AskToUpdateLinks = $false
  $book = $excel.Workbooks.Open($sourcePath, 0, $false)
  # Freeze every previous Excel web-function BEFORE installing the explicit-only update button.
  # Old WEBSERVICE/FILTERXML formulas would otherwise re-fetch on ordinary F9.
  $sheet = $book.Worksheets.Item('NBP_дані')
  $sheet.Range('B5').Value2 = ''
  $sheet.Range('B6').Value2 = 0
  $sheet.Range('B7').Value2 = ''
  $sheet.Range('B8').Value2 = ''
  $rates = $book.Worksheets.Item('Курси_NBP')
  $rates.Range('G13:H52').ClearContents()
  $settings = $book.Worksheets.Item('Налаштування')
  $settings.Range('B11').Value2 = 'Вручну: кнопка Оновити з інтернету'
  $settings.Range('B15').Value2 = 'Автоматичні інтернет-запити вимкнено'
  # Lightweight administrator panel; no UserForm dependencies.
  $form = $book.Worksheets.Add([System.Type]::Missing, $book.Worksheets.Item('Зміст'))
  $form.Name = 'Форма_адміністратора'
  $form.Range('A1').Value2 = 'ТРАНСПОРТНИЙ ЦЕНТР / ВВЕДЕННЯ РЕЙСУ'
  $labels = @('Водій (назва вкладки)','Дата нарахування','Замовник','Маршрут','Дата виїзду','Дата прибуття','Одометр початковий','Одометр кінцевий','Дохід USD','Отримано USD','Кермування, год','Простій, год','Примітка','Втрати часу, год')
  for ($i=0; $i -lt $labels.Count; $i++) { $form.Cells.Item($i + 5,1).Value2 = $labels[$i] }
  $form.Range('B6').Value2 = [double]([datetime]::Today.ToOADate())
  $form.Range('B6').NumberFormat = 'dd.mm.yyyy'
  $form.Range('B9:B10').NumberFormat = 'dd.mm.yyyy'
  $form.Columns.Item('A').ColumnWidth = 29; $form.Columns.Item('B').ColumnWidth = 37
  $form.Range('A1:B1').Font.Bold = $true
  $form.Range('A1:B19').Font.Name = 'Times New Roman'
  $form.Range('A1:B19').Font.Size = 14
  $form.Range('B5:B18').Interior.Color = 15136253
  $book.Names.Add('TC_Drivers', '=Реєстр_водіїв!$A$13:$A$52') | Out-Null
  $form.Range('B5').Validation.Delete()
  $form.Range('B5').Validation.Add(3,1,1,'=TC_Drivers') # Excel drop-down: registered drivers
  $form.Range('D5').Value2 = 'Веб-дані оновлюються ТІЛЬКИ окремою кнопкою.'
  $form.Range('D5').Font.Bold = $true
  $form.Columns.Item('D').ColumnWidth = 60
  $audit = $book.Worksheets.Add([System.Type]::Missing, $book.Worksheets.Item($book.Worksheets.Count))
  $audit.Name = 'Журнал_операцій'
  $audit.Range('A1').Value2 = 'Час'; $audit.Range('B1').Value2 = 'Користувач'
  $audit.Range('C1').Value2 = 'Дія'; $audit.Range('D1').Value2 = 'ID рейсу'
  $audit.Range('A1:D1').Font.Bold = $true
  $audit.Range('A1:D1').Font.Name = 'Times New Roman'
  $audit.Range('A1:D1').Font.Size = 14
  $audit.Columns.Item('A').ColumnWidth = 23; $audit.Columns.Item('B').ColumnWidth = 20
  $audit.Columns.Item('C').ColumnWidth = 24; $audit.Columns.Item('D').ColumnWidth = 28
  # Trust access to VBA project object model is necessary only during installation.
  # Avoid legacy ANSI .bas Import: preserve Ukrainian VBA names via Unicode COM BSTR.
  $vbaText = [System.IO.File]::ReadAllText($modulePath, [System.Text.Encoding]::UTF8)
  $vbaText = $vbaText -replace '(?m)^Attribute VB_Name = "[^"]*"\r?\n', ''
  $vbcomp = $book.VBProject.VBComponents.Add(1) # vbext_ct_StdModule
  $vbcomp.Name = 'TransportCenter'
  $vbcomp.CodeModule.AddFromString($vbaText)
  function Add-TcButton($sheet, [string]$caption, [string]$macro, [int]$x, [int]$y) {
    $shape = $sheet.Shapes.AddShape(5,$x,$y,173,33)
    $shape.TextFrame.Characters().Text = $caption
    $shape.TextFrame.Characters().Font.Name = 'Times New Roman'
    $shape.TextFrame.Characters().Font.Size = 12
    $shape.OnAction = "'" + $Destination + "'!TransportCenter." + $macro
  }
  Add-TcButton $form 'Зберегти рейс' 'ЗберегтиДані' 440 140
  Add-TcButton $form 'На головну' 'ПерейтиДоЗвітів' 440 183
  $contents = $book.Worksheets.Item('Зміст')
  Add-TcButton $contents 'Заповнити дані' 'ЗаповнитиДані' 620 40
  Add-TcButton $contents 'Додати водія' 'ДодатиВодія' 620 83
  Add-TcButton $contents 'Звіти' 'ПерейтиДоЗвітів' 620 126
  Add-TcButton $contents 'Резервна копія' 'РезервнаКопія' 620 169
  Add-TcButton $contents 'Оновити з інтернету' 'ОновитиЗІнтернету' 620 212
  Add-TcButton $contents 'Закрити рейс' 'ЗакритиРейс' 620 255
  Add-TcButton $contents 'PDF-звіт' 'СформуватиЗвіт' 620 298
  Add-TcButton $contents 'Відновити копію' 'ВідновитиРезервнуКопію' 620 341
  # Keep the content index physically FIRST as required.
  if ($book.Worksheets.Item(1).Name -ne 'Зміст') { $contents.Move($book.Worksheets.Item(1)) }
  $excel.Calculation = -4105 # xlCalculationAutomatic: arithmetic works; web formulas are removed
  $book.SaveAs($destinationPath, 52)
  if ($book.Worksheets.Item(1).Name -ne 'Зміст') { throw 'Зміст не є першим аркушем!' }
  Write-Host "Створено: $destinationPath"
  Write-Host 'Автоматичні вебформули відключено. Інтернет-запити тільки після натискання кнопки.'
}
finally {
 if ($book -ne $null) { $book.Close($false) }
 if ($excel -ne $null) { $excel.Quit() }
 [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
