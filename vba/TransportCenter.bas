Attribute VB_Name = "TransportCenter"
Option Explicit
Private Const BACKUP_FOLDER As String = "_Резервні_копії"
Private Const FORM_NAME As String = "Форма_адміністратора"

Private Function Ws(ByVal nm As String) As Worksheet
    Set Ws = ThisWorkbook.Worksheets(nm)
End Function

Private Function BackupNow(ByVal label As String) As String
    Dim dirName As String, pathName As String
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise 1001, , "Спочатку збережіть книгу."
    dirName = ThisWorkbook.Path & Application.PathSeparator & BACKUP_FOLDER
    If Dir(dirName, vbDirectory) = "" Then MkDir dirName
    pathName = dirName & Application.PathSeparator & Format(Now, "yyyymmdd_hhnnss") & "_" & label & ".xlsm"
    If Len(Dir(pathName)) > 0 Then pathName = Replace(pathName, ".xlsm", "_" & CStr(CLng(Timer * 100)) & ".xlsm")
    ThisWorkbook.SaveCopyAs pathName
    BackupNow = pathName
End Function

Public Sub РезервнаКопія()
    On Error GoTo failed
    MsgBox "Створено: " & BackupNow("manual"), vbInformation
    Exit Sub
failed:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub ЗаповнитиДані()
    On Error GoTo failed
    With Ws(FORM_NAME)
        .Range("B6").Value = Date
        .Activate
        .Range("B5").Select
    End With
    Exit Sub
failed:
    MsgBox "Форму не встановлено: " & Err.Description, vbExclamation
End Sub

Public Sub ЗберегтиДані()
    On Error GoTo failed
    Dim f As Worksheet, d As Worksheet, a As Worksheet, driver As String, tripID As String
    Dim r As Long, q As Long, backupPath As String, oldEvents As Boolean
    Set f = Ws(FORM_NAME): Set a = Ws("Нарахування")
    driver = Trim$(CStr(f.Range("B5").Value))
    If Application.CountIf(Ws("Реєстр_водіїв").Range("A13:A52"), driver) <> 1 Then Err.Raise 1002, , "Оберіть зареєстрованого водія."
    Set d = Ws(driver)
    If Not IsDate(f.Range("B6").Value) Or Not IsDate(f.Range("B9").Value) Or Not IsDate(f.Range("B10").Value) Then Err.Raise 1003, , "Перевірте дати."
    If CDate(f.Range("B10").Value) < CDate(f.Range("B9").Value) Then Err.Raise 1004, , "Дата прибуття раніше виїзду."
    If Trim$(CStr(f.Range("B7").Value)) = "" Or Trim$(CStr(f.Range("B8").Value)) = "" Then Err.Raise 1005, , "Вкажіть клієнта та маршрут."
    If Not IsNumeric(f.Range("B11").Value) Or Not IsNumeric(f.Range("B12").Value) Then Err.Raise 1006, , "Перевірте одометр."
    If CDbl(f.Range("B12").Value) < CDbl(f.Range("B11").Value) Then Err.Raise 1007, , "Кінцевий одометр менший за початковий."
    If Not IsNumeric(f.Range("B13").Value) Or Not IsNumeric(f.Range("B14").Value) Then Err.Raise 1008, , "Вкажіть дохід та оплату."
    If CDbl(f.Range("B13").Value) < 0 Or CDbl(f.Range("B14").Value) < 0 Or CDbl(f.Range("B14").Value) > CDbl(f.Range("B13").Value) Then Err.Raise 1009, , "Некоректні суми."
    For r = 13 To 92
        If Len(CStr(d.Cells(r, 1).Value2)) = 0 Then Exit For
    Next r
    For q = 13 To 132
        If Len(CStr(a.Cells(q, 1).Value2)) = 0 Then Exit For
    Next q
    If r > 92 Or q > 132 Then Err.Raise 1010, , "Журнал заповнений."
    tripID = "TC-" & Format(Now, "yyyymmdd-hhnnss") & "-" & Format(r, "00")
    If Application.CountIf(a.Range("A13:A132"), tripID) > 0 Then Err.Raise 1011, , "Повторний ID; спробуйте за секунду."
    If MsgBox("Записати рейс " & tripID & "?" & vbCrLf & "Буде створена резервна копія.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    backupPath = BackupNow("before_trip")
    oldEvents = Application.EnableEvents: Application.EnableEvents = False
    On Error GoTo rollback
    d.Cells(r, 1).Value = tripID
    d.Cells(r, 2).Value = CDate(f.Range("B9").Value)
    d.Cells(r, 3).Value = CDate(f.Range("B10").Value)
    d.Cells(r, 4).Value = f.Range("B8").Value
    d.Cells(r, 5).Value = CDbl(f.Range("B11").Value)
    d.Cells(r, 6).Value = CDbl(f.Range("B12").Value)
    If IsNumeric(f.Range("B15").Value) Then d.Cells(r, 8).Value = CDbl(f.Range("B15").Value)
    If IsNumeric(f.Range("B16").Value) Then d.Cells(r, 9).Value = CDbl(f.Range("B16").Value)
    a.Cells(q, 1).Value = tripID
    a.Cells(q, 2).Value = driver
    a.Cells(q, 3).Value = f.Range("B7").Value
    a.Cells(q, 4).Value = CDate(f.Range("B6").Value)
    a.Cells(q, 5).Value = CDbl(f.Range("B13").Value)
    a.Cells(q, 6).Value = CDbl(f.Range("B14").Value)
    a.Cells(q, 11).Value = f.Range("B17").Value
    Ws("Журнал_операцій").Cells(Ws("Журнал_операцій").Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 4).Value = Array(Now, Environ$("USERNAME"), "CREATE_TRIP", tripID)
    Application.EnableEvents = oldEvents
    Application.Calculate
    MsgBox "Рейс внесено: " & tripID, vbInformation
    Exit Sub
rollback:
    Application.EnableEvents = oldEvents
    MsgBox "Помилка після резервування. Копія: " & backupPath & vbCrLf & Err.Description, vbCritical
    Exit Sub
failed:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub ДодатиВодія()
    On Error GoTo failed
    Dim nm As String, fullName As String, vehicle As String, sh As Worksheet, rg As Worksheet, r As Long
    nm = Trim$(InputBox("Ім'я вкладки (наприклад Водій_03):", "Новий водій"))
    If Len(nm) = 0 Then Exit Sub
    If Len(nm) > 31 Or InStr(nm, "/") Or InStr(nm, "\\") Or InStr(nm, "[") Or InStr(nm, "]") Or InStr(nm, "*") Or InStr(nm, "?") Or InStr(nm, ":") Then Err.Raise 1101, , "Недопустима назва."
    On Error Resume Next
    Set sh = Ws(nm)
    On Error GoTo failed
    If Not sh Is Nothing Then Err.Raise 1102, , "Така вкладка вже є."
    Set rg = Ws("Реєстр_водіїв")
    For r = 13 To 52
        If Trim$(CStr(rg.Cells(r, 1).Value)) = "" Then Exit For
    Next r
    If r > 52 Then Err.Raise 1103, , "Реєстр заповнений."
    fullName = InputBox("ПІБ водія:", "Новий водій")
    vehicle = InputBox("Автомобіль:", "Новий водій")
    If Len(Trim$(fullName)) = 0 Then Exit Sub
    If MsgBox("Створити " & nm & "?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Call BackupNow("before_driver")
    Ws("ШАБЛОН_ВОДІЯ").Copy Before:=Ws("КІНЕЦЬ_ВОДІЇВ")
    Set sh = ActiveSheet: sh.Name = nm
    rg.Cells(r, 1).Value = nm
    rg.Cells(r, 2).Value = fullName
    rg.Cells(r, 3).Value = vehicle
    MsgBox "Водія додано. Реєстр і зміст оновляться формулами.", vbInformation
    Exit Sub
failed:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub ПерейтиДоЗвітів()
    Ws("Звіти").Activate
End Sub

Public Sub ЗакритиРейс()
    MsgBox "Закриття рейсу потребує перевірки всіх витрат і документів. На цьому етапі операцію не виконано.", vbInformation
End Sub

' ЄДИНА ТОЧКА ІНТЕРНЕТ-ЗАПИТІВ. Не викликати з Workbook_Open/Calculate/Change.
Public Sub ОновитиЗІнтернету()
    On Error GoTo failed
    Dim req As Object, doc As Object, node As Object
    Dim rateText As String, rateDate As String, rate As Double, backupPath As String
    If MsgBox("Оновити курс USD/PLN з NBP вручну?" & vbCrLf & "Паркінги без окремого перевіреного API не змінюватимуться.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Set req = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    req.setTimeouts 5000, 5000, 10000, 10000
    req.Open "GET", "https://api.nbp.pl/api/exchangerates/rates/a/usd/?format=xml", False
    req.Send
    If req.Status <> 200 Then Err.Raise 1201, , "NBP HTTP " & req.Status
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    If Not doc.LoadXML(req.responseText) Then Err.Raise 1202, , "Невірна XML-відповідь NBP."
    Set node = doc.SelectSingleNode("//*[local-name()='Mid']")
    If node Is Nothing Then Err.Raise 1203, , "У відповіді немає Mid."
    rateText = node.Text
    rate = Val(Replace(rateText, ",", "."))
    If rate < 1 Or rate > 10 Then Err.Raise 1204, , "Нетиповий USD/PLN; нічого не змінено."
    Set node = doc.SelectSingleNode("//*[local-name()='EffectiveDate']")
    If node Is Nothing Then Err.Raise 1205, , "Немає дати NBP."
    rateDate = node.Text
    If Not IsDate(rateDate) Then Err.Raise 1206, , "Некоректна дата NBP."
    backupPath = BackupNow("before_online")
    Ws("NBP_дані").Range("B6").Value = rate
    Ws("NBP_дані").Range("B7").Value = rateDate
    Ws("Налаштування").Range("B15").Value = "NBP вручну: " & Format(Now, "dd.mm.yyyy hh:nn")
    Application.Calculate
    MsgBox "Курс " & Format(rate, "0.0000") & " PLN/USD станом на " & rateDate & vbCrLf & "Паркінги не змінено.", vbInformation
    Exit Sub
failed:
    MsgBox "Онлайн-оновлення не виконано. Попередній курс збережено." & vbCrLf & Err.Description, vbExclamation
End Sub
