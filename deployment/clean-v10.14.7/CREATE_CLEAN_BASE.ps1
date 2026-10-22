#Requires -Version 5.1
[CmdletBinding()]
param([string]$OutputFile = (Join-Path $PSScriptRoot 'Транспортний_центр_v8_робочий_Times_New_Roman.xlsx'))
$ErrorActionPreference = 'Stop'
if (Test-Path -LiteralPath $OutputFile) {
    Write-Host "Clean base already exists: $OutputFile"
    exit 0
}
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Add(-4167) # xlWBATWorksheet
    $index = $book.Worksheets.Item(1)
    $index.Name = 'Зміст'
    $sheets = @(
        'Нарахування','NBP_дані','Курси_NBP','Налаштування',
        'Інструкція','Паркінги','Реєстр_водіїв','ШАБЛОН_ВОДІЯ',
        'КІНЕЦЬ_ВОДІЇВ','Звіти'
    )
    foreach ($name in $sheets) {
        $newSheet = $book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count))
        $newSheet.Name = $name
    }
    $index.Range('A1').Value2 = 'ТРАНСПОРТНИЙ ЦЕНТР / ЧИСТА ПОЧАТКОВА КНИГА'
    $index.Range('A3').Value2 = 'Порожня облікова база: додайте автопарк, водіїв та перший рейс.'
    $index.Columns.Item('A').ColumnWidth = 50
    $index.Range('A1:A3').Font.Name = 'Times New Roman'
    $index.Range('A1').Font.Size = 19
    $index.Range('A1').Font.Bold = $true

    $bill = $book.Worksheets.Item('Нарахування')
    $bh = @('Рейс ID','Водій','Замовник','Дата','Нараховано USD',
            'Оплачено USD','Борг USD','Факт км','Витрати USD','Контроль','Примітка')
    for($i=0;$i -lt $bh.Count;$i++){$bill.Cells.Item(12,$i+1).Value2=$bh[$i]}
    $bill.Range('A12:U12').Font.Bold=$true
    $bill.Columns.Item('A').ColumnWidth=24
    $bill.Columns.Item('B').ColumnWidth=21
    $bill.Columns.Item('C').ColumnWidth=25

    $book.Worksheets.Item('Налаштування').Range('B11').Value2='Курс валют установлюється вручну або кнопкою NBP'
    $book.Worksheets.Item('Інструкція').Range('A1').Value2='Транспортний центр — чиста початкова база.'
    $book.Worksheets.Item('Курси_NBP').Range('A12').Value2='Історію курсу не завантажено'
    $book.Worksheets.Item('NBP_дані').Range('A4').Value2='Курси ще не завантажено'

    $reg = $book.Worksheets.Item('Реєстр_водіїв')
    $reg.Range('A12').Value2='Водій / ID'
    $reg.Range('B12').Value2='Стан'

    $template = $book.Worksheets.Item('ШАБЛОН_ВОДІЯ')
    $template.Range('B2').Value2='ЖУРНАЛ ВОДІЯ'
    $dh = @('Рейс ID','Виїзд','Прибуття','Маршрут','Одометр поч.',
            'Одометр кін.','Пробіг','Кермування','Простій','Втрати часу',
            'Витрати','Дохід','Отримано','Примітки','Контроль')
    for($i=0;$i -lt $dh.Count;$i++){$template.Cells.Item(12,$i+1).Value2=$dh[$i]}
    $template.Range('A12:O12').Font.Bold = $true
    $template.Columns.Item('A').ColumnWidth = 27
    $park = $book.Worksheets.Item('Паркінги')
    $park.Cells.Item(10,1).Value2='Код'
    $park.Cells.Item(10,2).Value2='Назва'
    $park.Cells.Item(10,3).Value2='Тариф PLN'
    $book.Worksheets.Item('Звіти').Range('A1').Value2='Звіти формує Транспортний центр'
    $index.Activate()
    $book.SaveAs($OutputFile,51) # xlOpenXMLWorkbook; not a macro workbook
    $book.Close($false)
    $book=$null
    $excel.Quit()
    $excel=$null
    Write-Host "CLEAN base XLSX created: $OutputFile" -ForegroundColor Green
} catch {
    Write-Host ('CLEAN_BASE_ERROR: ' + $_.Exception.Message) -ForegroundColor Red
    throw
} finally {
    if($book -ne $null){try{$book.Close($false)}catch{}}
    if($excel -ne $null){try{$excel.Quit()}catch{}}
}
