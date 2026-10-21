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
    Dim f As Worksheet
    Set f = Ws(FORM_NAME)
    f.Range("B5:B18").ClearContents
    f.Range("B6").Value = Date
    f.Activate
    f.Range("B5").Select
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
    If IsNumeric(f.Range("B18").Value) Then d.Cells(r, 10).Value = CDbl(f.Range("B18").Value)
    Ws("Журнал_операцій").Cells(Ws("Журнал_операцій").Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 4).Value = Array(Now, Environ$("USERNAME"), "CREATE_TRIP", tripID)
    Application.EnableEvents = oldEvents
    Application.Calculate
    ThisWorkbook.Save
    MsgBox "Рейс внесено й збережено: " & tripID, vbInformation
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
    sh.Range("B5").Value = fullName
    sh.Range("E5").Value = vehicle
    rg.Cells(r, 1).Value = nm  ' B/C contain existing INDIRECT formulas; preserve them
    ThisWorkbook.Save
    MsgBox "Водія додано. Реєстр і зміст оновляться формулами.", vbInformation
    Exit Sub
failed:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub ПерейтиДоЗвітів()
    Ws("Звіти").Activate
End Sub

Public Sub ЗакритиРейс()
    On Error GoTo failed
    Dim id As String, a As Worksheet, d As Worksheet, q As Long, i As Long, hasCost As Boolean
    Set a = Ws("Нарахування")
    id = Trim$(InputBox("ID рейсу для закриття:", "Закрити рейс"))
    If Len(id) = 0 Then Exit Sub
    For q = 13 To 132
        If StrComp(Trim$(CStr(a.Cells(q, 1).Value2)), id, vbTextCompare) = 0 Then Exit For
    Next q
    If q > 132 Then Err.Raise 1301, , "Рейс не знайдено у нарахуваннях."
    If UCase$(Trim$(CStr(a.Cells(q, 12).Value2))) = "ЗАКРИТО" Then Err.Raise 1302, , "Рейс уже закрито."
    Application.Calculate
    If Trim$(CStr(a.Cells(q, 10).Value)) <> "OK" Then Err.Raise 1303, , "Спочатку усуньте помилки нарахування."
    Set d = Ws(CStr(a.Cells(q, 2).Value))
    For i = 102 To 201
        If StrComp(Trim$(CStr(d.Cells(i, 1).Value2)), id, vbTextCompare) = 0 Then
            hasCost = True
            If Trim$(CStr(d.Cells(i, 15).Value)) <> "OK" Then Err.Raise 1304, , "Є помилкові витрати, рядок " & i & "."
        End If
    Next i
    If MsgBox("Закрити рейс " & id & "?" & vbCrLf & "Витратні операції: " & IIf(hasCost, "є", "немає"), vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    Call BackupNow("before_close")
    a.Cells(q, 12).Value = "ЗАКРИТО"
    a.Cells(q, 13).Value = Now
    a.Cells(q, 14).Value = Environ$("USERNAME")
    With Ws("Журнал_операцій")
        .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 4).Value = Array(Now, Environ$("USERNAME"), "CLOSE_TRIP", id)
    End With
    ThisWorkbook.Save
    MsgBox "Рейс закрито. Резервну копію збережено.", vbInformation
    Exit Sub
failed:
    MsgBox Err.Description, vbExclamation
End Sub

' ЄДИНА ТОЧКА ІНТЕРНЕТ-ЗАПИТІВ. Не викликати з Workbook_Open/Calculate/Change.

Public Sub СформуватиЗвіт()
    On Error GoTo failed
    Dim folder As String, pathName As String
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise 1401, , "Спочатку збережіть книгу."
    folder = ThisWorkbook.Path & Application.PathSeparator & "_Звіти"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    pathName = folder & Application.PathSeparator & "Звіт_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    Ws("Звіти").ExportAsFixedFormat Type:=0, Filename:=pathName, Quality:=0, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    MsgBox "PDF-звіт сформовано:" & vbCrLf & pathName, vbInformation
    Exit Sub
failed:
    MsgBox "Не вдалося створити звіт: " & Err.Description, vbExclamation
End Sub

Public Sub ВідновитиРезервнуКопію()
    On Error GoTo failed
    Dim folder As String
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise 1501, , "Спочатку збережіть книгу."
    folder = ThisWorkbook.Path & Application.PathSeparator & BACKUP_FOLDER
    If Dir(folder, vbDirectory) = "" Then Err.Raise 1502, , "Резервних копій поки немає."
    Shell "explorer.exe " & Chr$(34) & folder & Chr$(34), vbNormalFocus
    MsgBox "Для відновлення закрийте Excel і запустіть RESTORE_BACKUP.cmd з папки програми." & vbCrLf & _
        "Поточна версія буде збережена перед відновленням.", vbInformation
    Exit Sub
failed:
    MsgBox Err.Description, vbExclamation
End Sub

' ONLY explicit button launches a web request.
Public Sub ОновитиЗІнтернету()
    On Error GoTo failed
    Dim req As Object, doc As Object, node As Object, rate As Double
    Dim effectiveDate As Date, stamp As String, hist As Worksheet, i As Long, target As Long
    If MsgBox("Отримати курс USD/PLN з NBP?" & vbCrLf & _
        "Автоматичного оновлення немає. Ціни паркінгів не зміняться.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Set req = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    req.setTimeouts 5000, 5000, 10000, 10000
    req.Open "GET", "https://api.nbp.pl/api/exchangerates/rates/a/usd/?format=xml", False
    req.Send
    If req.Status <> 200 Then Err.Raise 1201, , "NBP HTTP " & req.Status
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    If Not doc.LoadXML(req.responseText) Then Err.Raise 1202, , "Невірна XML-відповідь NBP."
    Set node = doc.SelectSingleNode("//*[local-name()='Mid']")
    If node Is Nothing Then Err.Raise 1203, , "Немає курсу Mid."
    rate = Val(Replace(node.Text, ",", "."))
    If rate < 1 Or rate > 10 Then Err.Raise 1204, , "Нетиповий курс: зміни скасовано."
    Set node = doc.SelectSingleNode("//*[local-name()='EffectiveDate']")
    If node Is Nothing Then Err.Raise 1205, , "Немає дати NBP."
    stamp = Trim$(node.Text)
    If Len(stamp) <> 10 Or Mid$(stamp, 5, 1) <> "-" Or Mid$(stamp, 8, 1) <> "-" Then Err.Raise 1206, , "Помилковий формат дати NBP."
    effectiveDate = DateSerial(CInt(Left$(stamp, 4)), CInt(Mid$(stamp, 6, 2)), CInt(Right$(stamp, 2)))
    If Format$(effectiveDate, "yyyy-mm-dd") <> stamp Then Err.Raise 1207, , "Неправильна дата NBP."
    Set hist = Ws("Курси_NBP")
    target = 0
    For i = 13 To 52
        If IsNumeric(hist.Cells(i, 1).Value2) And Len(CStr(hist.Cells(i, 1).Value2)) > 0 Then
            If CDbl(hist.Cells(i, 1).Value2) = CDbl(effectiveDate) Then
                target = -1
                If Abs(CDbl(hist.Cells(i, 2).Value2) - rate) > 0.000001 Then Err.Raise 1208, , "Котирування за цю дату відрізняється. Перевірте історію вручну."
                Exit For
            End If
        ElseIf target = 0 Then
            target = i
        End If
    Next i
    If target = 0 Then Err.Raise 1209, , "Історія курсів заповнена. Дані не змінено."
    Call BackupNow("before_online")
    If target > 0 Then
        hist.Cells(target, 1).Value = effectiveDate
        hist.Cells(target, 2).Value = rate
        hist.Cells(target, 5).Value = "NBP • ручне оновлення"
        hist.Cells(target, 6).Value = "https://api.nbp.pl/"
    End If
    Ws("NBP_дані").Range("B6").Value = rate
    Ws("NBP_дані").Range("B7").Value = stamp
    Ws("Налаштування").Range("B15").Value = "NBP вручну: " & Format(Now, "dd.mm.yyyy hh:nn")
    Ws("Налаштування").Calculate
    ThisWorkbook.Save
    MsgBox "Курс " & Format(rate, "0.0000") & " PLN/USD, дата " & stamp & vbCrLf & _
        "Паркінги не оновлювалися.", vbInformation
    Exit Sub
failed:
    MsgBox "Онлайн-оновлення не виконано. Попередній курс не змінено." & vbCrLf & Err.Description, vbExclamation
End Sub
