#Requires -Version 5.1
[CmdletBinding()]
param([string]$OutputFile = (Join-Path $PSScriptRoot 'Транспортний_центр_v8_робочий_Times_New_Roman.xlsx'))
$ErrorActionPreference = 'Stop'
if (Test-Path -LiteralPath $OutputFile) {
    Write-Host "Demo base exists, preserving it: $OutputFile"
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
        'Водій_01','Водій_02','КІНЕЦЬ_ВОДІЇВ','Звіти'
    )
    foreach ($name in $sheets) {
        $newSheet = $book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count))
        $newSheet.Name = $name
    }
    $index.Range('A1').Value2 = 'ДЕМО / БАЗОВА ТЕСТОВА КНИГА'
    $index.Range('A3').Value2 = 'Ця книга згенерована автоматично. Усі водії та перевезення вигадані.'
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

    $book.Worksheets.Item('Налаштування').Range('B8').Formula='=4'
    $book.Worksheets.Item('Налаштування').Range('B11').Value2='Тільки демонстраційні дані'
    $book.Worksheets.Item('Інструкція').Range('A1').Value2='ТЕСТОВА КНИГА. Використовувати лише для перевірки.'
    $book.Worksheets.Item('Курси_NBP').Range('A12').Value2='Історію курсу не завантажено'
    $book.Worksheets.Item('NBP_дані').Range('A4').Value2='Демо / без мережевих оновлень'

    $reg = $book.Worksheets.Item('Реєстр_водіїв')
    $reg.Range('A12').Value2='Водій / ID'
    $reg.Range('B12').Value2='Стан'
    $reg.Range('A13').Value2='Водій_01'
    $reg.Range('A14').Value2='Водій_02'

    $template = $book.Worksheets.Item('ШАБЛОН_ВОДІЯ')
    $template.Range('B2').Value2='ЖУРНАЛ ВОДІЯ / ДЕМО'
    $dh = @('Рейс ID','Виїзд','Прибуття','Маршрут','Одометр поч.',
            'Одометр кін.','Пробіг','Кермування','Простій','Втрати часу',
            'Витрати','Дохід','Отримано','Примітки','Контроль')
    for($i=0;$i -lt $dh.Count;$i++){$template.Cells.Item(12,$i+1).Value2=$dh[$i]}
    $template.Range('A12:O12').Font.Bold = $true
    $template.Columns.Item('A').ColumnWidth = 27
    for($i=1;$i -le 2;$i++){
        $name = ('Водій_{0:d2}' -f $i)
        $sheet = $book.Worksheets.Item($name)
        $template.Range('A1:O93').Copy($sheet.Range('A1'))
        $sheet.Range('B5').Value2='ДЕМОНСТРАЦІЙНИЙ ВОДІЙ'
    }
    $park = $book.Worksheets.Item('Паркінги')
    $park.Cells.Item(10,1).Value2='Код'
    $park.Cells.Item(10,2).Value2='Тестова назва'
    $park.Cells.Item(10,3).Value2='Тариф PLN'
    $park.Cells.Item(11,1).Value2='DEMO-PARK-01'
    $park.Cells.Item(11,2).Value2='Тестова стоянка A'
    $park.Cells.Item(11,3).Formula='=90'
    $park.Cells.Item(12,1).Value2='DEMO-PARK-02'
    $park.Cells.Item(12,2).Value2='Тестова стоянка B'
    $park.Cells.Item(12,3).Formula='=120'
    $book.Worksheets.Item('Звіти').Range('A1').Value2='Звіти формує Транспортний центр'
    $index.Activate()
    $book.SaveAs($OutputFile,51) # xlOpenXMLWorkbook; not a macro workbook
    $book.Close($false)
    $book=$null
    $excel.Quit()
    $excel=$null
    Write-Host "DEMO-only base XLSX created: $OutputFile" -ForegroundColor Green
} catch {
    Write-Host ('DEMO_BASE_ERROR: ' + $_.Exception.Message) -ForegroundColor Red
    throw
} finally {
    if($book -ne $null){try{$book.Close($false)}catch{}}
    if($excel -ne $null){try{$excel.Quit()}catch{}}
}
