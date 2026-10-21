# Catalog worksheets and lookup formulas (Excel COM, Windows; no Internet requests).
function New-TCCatalogSheet {
 param($book,[string]$name,[string[]]$headers,[string]$lastColumn,[int]$lastRow)
 $sh=$book.Worksheets.Add([System.Type]::Missing,$book.Worksheets.Item('Звіти'))
 $sh.Name=$name
 $sh.Range('A1').Value2='ТРАНСПОРТНИЙ ЦЕНТР / '+$name
 $sh.Range('A1').Font.Bold=$true
 $sh.Range('A3').Value2='Жовті поля — ручне заповнення довідника. Інтернет-запити автоматично не виконуються.'
 $sh.Range("A5:${lastColumn}5").Font.Bold=$true
 $sh.Range("A5:${lastColumn}5").Interior.Color=4210752
 $sh.Range("A5:${lastColumn}5").Font.Color=16777215
 $sh.Range("A1:${lastColumn}${lastRow}").Font.Name='Times New Roman'
 $sh.Range("A1:${lastColumn}${lastRow}").Font.Size=14
 for($i=0;$i -lt $headers.Length;$i++){ $sh.Cells.Item(5,$i+1).Value2=$headers[$i] }
 $sh.Range("A6:${lastColumn}${lastRow}").Interior.Color=15136253
 $sh.Rows.Item(5).RowHeight=38
 $sh.Rows.Item(5).WrapText=$true
 $sh.Activate()
 $excel=$book.Application
 $excel.ActiveWindow.SplitRow=5
 $excel.ActiveWindow.FreezePanes=$true
 return $sh
}
function Initialize-TCCatalogs {
 param($book,$form)
 $carHeaders=@('Код авто','Марка','Модель','Держ. номер','Пальне','Л/100 км','Пальне PLN/л','Бак, л','Статус','Примітка')
 $cars=New-TCCatalogSheet $book 'Автопарк' $carHeaders 'J' 65
 $cars.Range('A3').Value2='Для кожного автомобіля один сталий код. Ціни та норми задає адміністратор; чутливі дані залишаються локально.'
 $cars.Columns.Item('A').ColumnWidth=18
 $cars.Columns.Item('B').ColumnWidth=22
 $cars.Columns.Item('C').ColumnWidth=20
 $cars.Columns.Item('D').ColumnWidth=22
 $cars.Columns.Item('E').ColumnWidth=18
 $cars.Columns.Item('F').ColumnWidth=14
 $cars.Columns.Item('G').ColumnWidth=19
 $cars.Columns.Item('H').ColumnWidth=13
 $cars.Columns.Item('I').ColumnWidth=18
 $cars.Columns.Item('J').ColumnWidth=34
 $cars.Range('F6:H65').NumberFormat='0.00'
 $cars.Range('E6:E65').Validation.Add(3,1,1,'Дизель,Бензин,Газ,Електро,Інше')
 $cars.Range('I6:I65').Validation.Add(3,1,1,'Активне,Ремонт,Резерв,Неактивне')
 $book.Names.Add('TC_Cars','=Автопарк!$A$6:$A$65') | Out-Null

 $peopleHeaders=@('Аркуш водія','Прізвище','Ім''я','Телефон','Код авто','E-mail','Статус','ПІБ (авто)')
 $people=New-TCCatalogSheet $book 'Контакти_водіїв' $peopleHeaders 'H' 65
 $people.Range('A3').Value2='Персональні контакти зберігайте ЛИШЕ в робочому файлі, ніколи в публічному GitHub.'
 foreach($col in @('A','B','C','D','E','F','G','H')){ $people.Columns.Item($col).ColumnWidth=22 }
 $people.Columns.Item('D').ColumnWidth=24
 $people.Columns.Item('F').ColumnWidth=32
 $people.Columns.Item('H').ColumnWidth=32
 $people.Range('D6:D65').NumberFormat='@'
 $people.Range('E6:E65').Validation.Add(3,1,1,'=TC_Cars')
 $people.Range('G6:G65').Validation.Add(3,1,1,'Активний,Відпустка,Неактивний')
 $people.Range('H6').Formula='=IF(A6="","",IF(AND(B6="",C6=""),A6,TRIM(B6&" "&C6)&" ["&A6&"]"))'
 $people.Range('H6:H65').FillDown()
 $book.Names.Add('TC_DriverDisplay','=Контакти_водіїв!$H$6:$H$65') | Out-Null
 $people.Range('H6:H65').Interior.Color=14277081
 # Preserve existing two driver tabs: placeholder IDs, no invented personal details.
 $people.Range('A6').Value2='Водій_01'
 $people.Range('A7').Value2='Водій_02'
 $people.Range('G6:G7').Value2='Активний'

 $routeHeaders=@('ID маршруту','Початок','Призначення','Км затверджені','Час, год','Дороги/країни','Усіх рейсів','Км факт','Контроль','Примітка','Закрито рейсів')
 $routes=New-TCCatalogSheet $book 'Маршрути_Європа' $routeHeaders 'K' 65
 $routes.Range('A3').Value2='Км не підставляємо навмання: адміністратор вводить відстань за погодженим маршрутом. Рейси рахує Excel.'
 foreach($col in @('A','B','C')){ $routes.Columns.Item($col).ColumnWidth=21 }
 $routes.Columns.Item('D').ColumnWidth=19
 $routes.Columns.Item('E').ColumnWidth=15
 $routes.Columns.Item('F').ColumnWidth=26
 $routes.Columns.Item('G').ColumnWidth=16
 $routes.Columns.Item('H').ColumnWidth=16
 $routes.Columns.Item('I').ColumnWidth=21
 $routes.Columns.Item('J').ColumnWidth=36
 $routes.Columns.Item('K').ColumnWidth=19
 $routes.Range('D6:E65').NumberFormat='0.0'
 $seed=@(
  @('WAW-BER','Варшава','Берлін'),
  @('WAW-PRG','Варшава','Прага'),
  @('WAW-VIE','Варшава','Відень'),
  @('WAW-HAM','Варшава','Гамбург'),
  @('WAW-AMS','Варшава','Амстердам'),
  @('WAW-RTM','Варшава','Роттердам'),
  @('WRO-DRE','Вроцлав','Дрезден'),
  @('POZ-BER','Познань','Берлін'),
  @('KRK-BUD','Краків','Будапешт'),
  @('GDN-HAM','Гданськ','Гамбург'),
  @('WAW-MIL','Варшава','Мілан'),
  @('PRG-MUN','Прага','Мюнхен')
 )
 for($i=0;$i -lt $seed.Count;$i++){
  $rr=$i+6
  $routes.Cells.Item($rr,1).Value2=$seed[$i][0]
  $routes.Cells.Item($rr,2).Value2=$seed[$i][1]
  $routes.Cells.Item($rr,3).Value2=$seed[$i][2]
  $routes.Cells.Item($rr,10).Value2='Шаблон · узгодити кілометраж'
 }
 $routes.Range('G6').Formula='=IF(A6="","",COUNTIF(Нарахування!$O$13:$O$132,A6))'
 $routes.Range('H6').Formula='=IF(A6="","",SUMIF(Нарахування!$O$13:$O$132,A6,Нарахування!$R$13:$R$132))'
 $routes.Range('I6').Formula='=IF(A6="","",IF(D6>0,"ГОТОВО","НЕМАЄ КМ"))'
 $routes.Range('K6').Formula='=IF(A6="","",COUNTIFS(Нарахування!$O$13:$O$132,A6,Нарахування!$L$13:$L$132,"ЗАКРИТО"))'
 $routes.Range('G6:G65').FillDown()
 $routes.Range('H6:H65').FillDown()
 $routes.Range('I6:I65').FillDown()
 $routes.Range('K6:K65').FillDown()
 $routes.Range('G6:I65').Interior.Color=14277081
 $routes.Range('K6:K65').Interior.Color=14277081
 $book.Names.Add('TC_Routes','=Маршрути_Європа!$A$6:$A$65') | Out-Null

 $profile=$book.Worksheets.Add([System.Type]::Missing,$book.Worksheets.Item('Звіти'))
 $profile.Name='Картка_водія'
 $profile.Range('A1').Value2='НОВИЙ ВОДІЙ / КАРТКА'
 $names=@('Прізвище','Ім''я','Телефон','E-mail','Автомобіль (код)','Назва вкладки (необов''язково)')
 for($i=0;$i -lt $names.Count;$i++){ $profile.Cells.Item($i+5,1).Value2=$names[$i] }
 $profile.Columns.Item('A').ColumnWidth=39
 $profile.Columns.Item('B').ColumnWidth=36
 $profile.Range('A1:B12').Font.Name='Times New Roman'
 $profile.Range('A1:B12').Font.Size=14
 $profile.Range('A1').Font.Bold=$true
 $profile.Range('B5:B10').Interior.Color=15136253
 $profile.Range('B7').NumberFormat='@'
 $profile.Range('B9').Validation.Add(3,1,1,'=TC_Cars')
 $profile.Range('A12').Value2='Заповни картку один раз; надалі водій та авто обираються зі списків.'

 # Administrator entry form: route+vehicle as dropdowns; expense forecast is formula-only.
 $form.Range('A8').Value2='Маршрут (ID зі списку)'
 $form.Range('A19').Value2='Інше авто (необов''язково)'
 $form.Range('A20').Value2='Ціна пального PLN/л (заміна)'
 $form.Range('A21').Value2='Пробіг план, км (заміна)'
 $form.Range('A22').Value2='Резерв на паливо, %'
 $form.Range('B22').Value2=0.1
 $form.Range('B22').NumberFormat='0%'
 $form.Range('B19').Validation.Add(3,1,1,'=TC_Cars')
 $form.Range('B8').Validation.Delete()
 $form.Range('B8').Validation.Add(3,1,1,'=TC_Routes')
 $form.Range('B19:B22').Interior.Color=15136253
 $form.Range('A19:B29').Font.Name='Times New Roman'
 $form.Range('A19:B29').Font.Size=14
 $form.Range('A19:A22').Font.Bold=$true
 $form.Range('A19').EntireRow.RowHeight=24
 $form.Range('E5').Value2='ПРОГНОЗ ПАЛЬНОГО'
 $form.Range('E5').Font.Bold=$true
 $calcLabels=@('Водій ПІБ','Авто (ID)','Марка / номер','Норма, л/100 км','Маршрут, км','Ціна PLN/л','Прогноз, літри','Прогноз, PLN','Орієнтовно USD','Статус прогнозу','Виконано рейсів','Маршрут (міста)')
 for($i=0;$i -lt $calcLabels.Count;$i++){$form.Cells.Item($i+7,5).Value2=$calcLabels[$i]}
 $form.Range('F7').Formula='=IF($B$5="","",$B$5)'
 $form.Range('F8').Formula='=IF($B$19<>"",$B$19,IF($B$5="","",IFERROR(IF(INDEX(Контакти_водіїв!$E$6:$E$65,MATCH($B$5,Контакти_водіїв!$H$6:$H$65,0))=0,"",INDEX(Контакти_водіїв!$E$6:$E$65,MATCH($B$5,Контакти_водіїв!$A$6:$A$65,0))),"")))'
 $form.Range('F9').Formula='=IF(F8="","",IFERROR(INDEX(Автопарк!$B$6:$B$65,MATCH(F8,Автопарк!$A$6:$A$65,0))&" "&INDEX(Автопарк!$C$6:$C$65,MATCH(F8,Автопарк!$A$6:$A$65,0))&" / "&INDEX(Автопарк!$D$6:$D$65,MATCH(F8,Автопарк!$A$6:$A$65,0)),""))'
 $form.Range('F10').Formula='=IF(F8="","",IFERROR(INDEX(Автопарк!$F$6:$F$65,MATCH(F8,Автопарк!$A$6:$A$65,0)),0))'
 $form.Range('F11').Formula='=IF($B$21>0,$B$21,IFERROR(INDEX(Маршрути_Європа!$D$6:$D$65,MATCH($B$8,Маршрути_Європа!$A$6:$A$65,0)),0))'
 $form.Range('F12').Formula='=IF($B$20>0,$B$20,IFERROR(INDEX(Автопарк!$G$6:$G$65,MATCH(F8,Автопарк!$A$6:$A$65,0)),0))'
 $form.Range('F13').Formula='=IF(OR(F10<=0,F11<=0),"",ROUND(F11*F10/100,2))'
 $form.Range('F14').Formula='=IF(OR(F13="",F12<=0),"",ROUND(F13*F12*(1+$B$22),2))'
 $form.Range('F15').Formula='=IF(F14="","",IFERROR(IF(Налаштування!$B$8>0,F14/Налаштування!$B$8,""),""))'
 $form.Range('F16').Formula='=IF(AND(F8<>"",F10>0,F11>0,F12>0),"ГОТОВО","НЕПОВНО: авто / норма / км / ціна")'
 $form.Range('F17').Formula='=IF($B$8="","",COUNTIF(Нарахування!$O$13:$O$132,$B$8))'
 $form.Range('F18').Formula='=IF($B$8="","",IFERROR(INDEX(Маршрути_Європа!$B$6:$B$65,MATCH($B$8,Маршрути_Європа!$A$6:$A$65,0))&" - "&INDEX(Маршрути_Європа!$C$6:$C$65,MATCH($B$8,Маршрути_Європа!$A$6:$A$65,0)),""))'
 $form.Range('E5:F18').Font.Name='Times New Roman'
 $form.Range('E5:F18').Font.Size=14
 $form.Columns.Item('E').ColumnWidth=25
 $form.Columns.Item('F').ColumnWidth=45
 $form.Range('F10:F15').NumberFormat='0.00'
 $form.Range('F7:F18').Interior.Color=14277081
 $form.Range('E20').Value2='Фінансова оцінка: тариф пального задає адміністратор. Поточних цін з інтернету не отримуємо.'

 # Fixed forecast metadata per trip, never dependent on later tariff/rate changes.
 $a=$book.Worksheets.Item('Нарахування')
 $forecastHeaders=@('Маршрут (ID)','Авто (ID)','Паливо прогноз PLN','Км фактичні','Літри прогноз','Ціна PLN/л','Км планові')
 for($i=0;$i -lt $forecastHeaders.Count;$i++){
   $a.Cells.Item(12,15+$i).Value2=$forecastHeaders[$i]
 }
 $a.Range('O12:U12').Font.Name='Times New Roman'
 $a.Range('O12:U12').Font.Size=14
 $a.Range('O12:U12').Font.Bold=$true
 $a.Range('O12:U12').Interior.Color=4210752
 $a.Range('O12:U12').Font.Color=16777215
 $a.Range('O13:U132').Interior.Color=14277081
 foreach($col in @('O','P','Q','R','S','T','U')){$a.Columns.Item($col).ColumnWidth=21}
 $a.Range('Q13:U132').NumberFormat='0.00'
}
