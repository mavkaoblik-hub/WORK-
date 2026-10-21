Attribute VB_Name = "TransportCatalogs"
Option Explicit

' Catalogs are local workbook data, not external services.
Private Function SheetExists(ByVal nm As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function

Private Function NextDriverName() As String
    Dim i As Long, candidate As String
    For i = 1 To 999
        candidate = "Водій_" & Format$(i, "00")
        If Not SheetExists(candidate) Then
            NextDriverName = candidate
            Exit Function
        End If
    Next i
    Err.Raise 1701, , "Вільних назв водіїв немає."
End Function

Private Function CarRow(ByVal carId As String) As Long
    Dim k As Long, c As Worksheet
    Set c = ThisWorkbook.Worksheets("Автопарк")
    For k = 6 To 65
        If StrComp(Trim$(CStr(c.Cells(k, 1).Value2)), carId, vbTextCompare) = 0 Then
            CarRow = k
            Exit Function
        End If
    Next k
End Function

Public Sub ВідкритиАвтопарк()
    ThisWorkbook.Worksheets("Автопарк").Activate
End Sub

Public Sub ВідкритиКонтакти()
    ThisWorkbook.Worksheets("Контакти_водіїв").Activate
End Sub

Public Sub ВідкритиМаршрути()
    ThisWorkbook.Worksheets("Маршрути_Європа").Activate
End Sub

Public Sub ЗареєструватиВодія()
    On Error GoTo failed
    Dim f As Worksheet, p As Worksheet, reg As Worksheet, sh As Worksheet
    Dim surname As String, firstName As String, telephone As String, email As String
    Dim nm As String, carId As String, plate As String, fullName As String
    Dim profileRow As Long, registerRow As Long, carIndex As Long
    Dim oldEvents As Boolean, back As String
    Set f = ThisWorkbook.Worksheets("Картка_водія")
    Set p = ThisWorkbook.Worksheets("Контакти_водіїв")
    Set reg = ThisWorkbook.Worksheets("Реєстр_водіїв")
    surname = Trim$(CStr(f.Range("B5").Value))
    firstName = Trim$(CStr(f.Range("B6").Value))
    telephone = Trim$(CStr(f.Range("B7").Text))
    email = Trim$(CStr(f.Range("B8").Value))
    carId = Trim$(CStr(f.Range("B9").Value))
    nm = Trim$(CStr(f.Range("B10").Value))
    If surname = "" Or firstName = "" Then Err.Raise 1702, , "Введіть прізвище та ім'я."
    If Len(surname) > 80 Or Len(firstName) > 80 Then Err.Raise 1703, , "Надто довге ім'я."
    If nm = "" Then nm = NextDriverName()
    If Len(nm) > 31 Or InStr(nm, "/") > 0 Or InStr(nm, "\") > 0 Or InStr(nm, ":") > 0 Or _
       InStr(nm, "[") > 0 Or InStr(nm, "]") > 0 Or InStr(nm, "*") > 0 Or InStr(nm, "?") > 0 Or _
       InStr(nm, "'") > 0 Then Err.Raise 1704, , "Недопустима назва вкладки."
    If SheetExists(nm) Then Err.Raise 1705, , "Назва вкладки вже зайнята."
    If carId <> "" Then
        carIndex = CarRow(carId)
        If carIndex = 0 Then Err.Raise 1706, , "Спочатку зареєструйте автомобіль в «Автопарк»."
        plate = Trim$(CStr(ThisWorkbook.Worksheets("Автопарк").Cells(carIndex, 4).Value2))
    End If
    For registerRow = 13 To 52
        If Trim$(CStr(reg.Cells(registerRow, 1).Value2)) = "" Then Exit For
    Next registerRow
    If registerRow > 52 Then Err.Raise 1707, , "Ліміт реєстру — 40 водіїв."
    For profileRow = 6 To 65
        If Trim$(CStr(p.Cells(profileRow, 1).Value2)) = "" Then Exit For
    Next profileRow
    If profileRow > 65 Then Err.Raise 1708, , "Ліміт карток — 60 водіїв."
    fullName = surname & " " & firstName
    If MsgBox("Додати водія: " & fullName & vbCrLf & "Автомобіль: " & _
       IIf(carId = "", "не призначено", carId) & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    back = BackupNow("before_new_driver")
    oldEvents = Application.EnableEvents
    On Error GoTo rollback
    Application.EnableEvents = False
    ThisWorkbook.Worksheets("ШАБЛОН_ВОДІЯ").Copy Before:=ThisWorkbook.Worksheets("КІНЕЦЬ_ВОДІЇВ")
    Set sh = ActiveSheet
    sh.Name = nm
    sh.Range("B5").Value = fullName
    sh.Range("E5").Value = plate
    reg.Cells(registerRow, 1).Value = nm
    p.Cells(profileRow, 1).Value = nm
    p.Cells(profileRow, 2).Value = surname
    p.Cells(profileRow, 3).Value = firstName
    p.Cells(profileRow, 4).NumberFormat = "@"
    p.Cells(profileRow, 4).Value = telephone
    p.Cells(profileRow, 5).Value = carId
    p.Cells(profileRow, 6).Value = email
    p.Cells(profileRow, 7).Value = "Активний"
    With ThisWorkbook.Worksheets("Журнал_операцій")
        .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Resize(1, 4).Value = _
            Array(Now, Environ$("USERNAME"), "NEW_DRIVER", nm)
    End With
    f.Range("B5:B10").ClearContents
    Application.EnableEvents = oldEvents
    ThisWorkbook.Save
    MsgBox "Створено " & nm & ". Контактна картка, реєстр і зміст оновлені.", vbInformation
    Exit Sub
rollback:
    Application.EnableEvents = oldEvents
    MsgBox "Помилка додавання водія: " & Err.Description & vbCrLf & _
           "Попередня версія: " & back, vbCritical
    Exit Sub
failed:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub СинхронізуватиВодіїв()
    On Error GoTo failed
    Dim p As Worksheet, cars As Worksheet, sh As Worksheet
    Dim k As Long, carId As String, carIndex As Long, nm As String
    Dim fullName As String, plate As String, changes As Long
    Set p = ThisWorkbook.Worksheets("Контакти_водіїв")
    Set cars = ThisWorkbook.Worksheets("Автопарк")
    If MsgBox("Оновити ПІБ та автомобілі на вкладках водіїв із довідників?", _
        vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    Call BackupNow("before_sync_drivers")
    For k = 6 To 65
        nm = Trim$(CStr(p.Cells(k, 1).Value2))
        If nm <> "" And SheetExists(nm) Then
            Set sh = ThisWorkbook.Worksheets(nm)
            fullName = Trim$(CStr(p.Cells(k, 2).Value2) & " " & CStr(p.Cells(k, 3).Value2))
            carId = Trim$(CStr(p.Cells(k, 5).Value2))
            If fullName <> "" Then sh.Range("B5").Value = fullName
            If carId <> "" Then
                carIndex = CarRow(carId)
                If carIndex > 0 Then
                    plate = Trim$(CStr(cars.Cells(carIndex, 4).Value2))
                    sh.Range("E5").Value = plate
                End If
            End If
            changes = changes + 1
        End If
    Next k
    ThisWorkbook.Save
    MsgBox "Синхронізовано картки водіїв: " & changes, vbInformation
    Exit Sub
failed:
    MsgBox "Синхронізацію не завершено: " & Err.Description, vbExclamation
End Sub
