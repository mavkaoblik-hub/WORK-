Attribute VB_Name = "TransportApp"
Option Explicit

' Transport Center v10.  The workbook is an OFFLINE application.
' All interaction with operational data happens through frmTCAdmin dialogs.
' No Workbook_Open / Worksheet_Change events or automatic HTTP requests.
Private mMode As String
Private mHostBook As Workbook ' persistent host workbook reference while frmTCAdmin is modal
Private mUiDrivers As Collection
Private mUiVehicles As Collection
Private mUiRoutes As Collection
Private mUiParkings As Collection
Private mUiTrips As Collection
Private mUiSummary As String
Private mUiLoading As Boolean
Private Const FIRST_DATA As Long = 6
Private Const MAX_MASTER_DRIVERS As Long = 40
Private Const MAX_TRIP_ROWS As Long = 50000

Private Function WS(ByVal sheetName As String) As Worksheet
    ' Resolve worksheet through the book captured BEFORE showing the modal form.
    ' Hidden/VeryHidden sheets are valid data sources and must never be activated.
    Dim candidate As Worksheet, bookName As String, visibleNames As String
    On Error GoTo ResolveFailed
    If mHostBook Is Nothing Then Set mHostBook = ThisWorkbook
    If mHostBook Is Nothing Then Err.Raise vbObjectError + 2401, , "Робочу книгу не визначено."
    bookName = mHostBook.Name
    For Each candidate In mHostBook.Worksheets
        If StrComp(candidate.Name, sheetName, vbTextCompare) = 0 Then
            Set WS = candidate
            Exit Function
        End If
        If Len(visibleNames) < 350 Then visibleNames = visibleNames & candidate.Name & "; "
    Next candidate
    Err.Raise vbObjectError + 2402, , "Аркуш '" & sheetName & "' відсутній. Книга: " & bookName & ". Є: " & visibleNames
    Exit Function
ResolveFailed:
    Dim details As String
    details = Err.Description
    Err.Raise vbObjectError + 2403, "TransportApp.WS", "Помилка доступу до '" & sheetName & "' (" & bookName & "): " & details
End Function

Private Function Tail(ByVal ws As Worksheet, ByVal col As Long, ByVal startAt As Long) As Long
    Dim k As Long
    k = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If k < startAt Then k = startAt - 1
    Tail = k + 1
End Function

Private Function UiText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    UiText = Trim$(CStr(v))
End Function

Private Function IdPrefix(ByVal s As String) As String
    ' ComboBox is intentionally empty for new items and may be empty for optional choices.
    ' Split(vbNullString)(0) raises runtime error 9 (Subscript out of range).
    ' Use a delimiter position instead of indexing a potentially empty array.
    Dim delimiterPos As Long
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    delimiterPos = InStr(1, s, " | ", vbBinaryCompare)
    If delimiterPos > 0 Then
        IdPrefix = Trim$(Left$(s, delimiterPos - 1))
    Else
        IdPrefix = s
    End If
End Function

Private Function DriverIdFromDisplay(ByVal s As String) As String
    Dim rowId As Long, p As Worksheet
    Set p = WS("Контакти_водіїв")
    For rowId = 6 To p.Cells(p.Rows.Count, 1).End(xlUp).Row
        If Len(Trim$(CStr(p.Cells(rowId, 1).Value2))) <> 0 Then
            If StrComp(Trim$(CStr(p.Cells(rowId, 8).Value)), Trim$(s), vbTextCompare) = 0 Then
                DriverIdFromDisplay = Trim$(CStr(p.Cells(rowId, 1).Value2))
                Exit Function
            End If
        End If
    Next rowId
End Function

Private Function MasterRow(ByVal sheetName As String, ByVal needle As String) As Long
    Dim p As Worksheet, i As Long
    Set p = WS(sheetName)
    For i = FIRST_DATA To p.Cells(p.Rows.Count, 1).End(xlUp).Row
        If StrComp(Trim$(CStr(p.Cells(i, 1).Value2)), needle, vbTextCompare) = 0 Then
            MasterRow = i
            Exit Function
        End If
    Next i
End Function

Private Function TripRow(ByVal tripId As String) As Long
    TripRow = MasterRow("TC_Рейси", tripId)
End Function

Private Function NonNeg(ByVal raw As String, ByVal caption As String, Optional ByVal allowEmpty As Boolean = False) As Double
    If Trim$(raw) = "" And allowEmpty Then Exit Function
    If Trim$(raw) = "" Or Not IsNumeric(raw) Then Err.Raise 2101, , "Введіть число: " & caption
    NonNeg = CDbl(raw)
    If NonNeg < 0 Then Err.Raise 2102, , "Від'ємне значення: " & caption
End Function

Private Function DateOrToday(ByVal raw As String) As Date
    Dim s As String
    s = Trim$(raw)
    If s = "" Then DateOrToday = Date: Exit Function
    If Len(s) = 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        DateOrToday = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Right$(s, 2)))
        If Format$(DateOrToday, "yyyy-mm-dd") <> s Then Err.Raise 2110, , "Невірна дата: " & s
        Exit Function
    End If
    If Not IsDate(s) Then Err.Raise 2111, , "Невірна дата: " & s
    DateOrToday = CDate(s)
End Function

Private Function BackupApp(ByVal tag As String) As String
    Dim folder As String, filename As String
    If ThisWorkbook.Path = "" Then Err.Raise 2112, , "Спочатку збережіть книгу."
    folder = ThisWorkbook.Path & Application.PathSeparator & "_Резервні_копії"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    filename = folder & Application.PathSeparator & Format$(Now, "yyyymmdd_hhnnss") & "_" & tag & ".xlsm"
    If Dir(filename) <> "" Then filename = Replace(filename, ".xlsm", "_" & Format$(Timer * 100, "0") & ".xlsm")
    ThisWorkbook.SaveCopyAs filename
    BackupApp = filename
End Function

Private Sub EventLog(ByVal operation As String, ByVal detail As String)
    Dim w As Worksheet, r As Long
    Set w = WS("TC_Події")
    r = Tail(w, 1, FIRST_DATA)
    w.Cells(r, 1).Value = Now
    w.Cells(r, 2).Value = Environ$("USERNAME")
    w.Cells(r, 3).Value = operation
    w.Cells(r, 4).Value = detail
End Sub

Public Sub TC_UI_NewTrip()
    ShowTC "NEW"
End Sub
Public Sub TC_UI_StartTrip()
    ShowTC "START"
End Sub
Public Sub TC_UI_FinishTrip()
    ShowTC "FINISH"
End Sub
Public Sub TC_UI_Expense()
    ShowTC "COST"
End Sub
Public Sub TC_UI_Fleet()
    ShowTC "CAR"
End Sub
Public Sub TC_UI_Driver()
    ShowTC "DRIVER"
End Sub
Public Sub TC_UI_Route()
    ShowTC "ROUTE"
End Sub
Public Sub TC_UI_Parking()
    ShowTC "PARKING"
End Sub
Public Sub TC_UI_Report()
    ShowTC "REPORT"
End Sub
Public Sub TC_UI_Backup()
    On Error GoTo Failed
    If MsgBox("Створити резервну копію поточної книги?", vbYesNo + vbQuestion, "Резервування") = vbYes Then
        MsgBox "Резервну копію створено:" & vbCrLf & BackupApp("manual"), vbInformation
    End If
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
Public Sub TC_UI_OnlineNBP()
    ' Existing manual-only NBP procedure; no events call this method.
    TransportCenter.TC_RefreshNBP
End Sub

Private Sub ShowTC(ByVal modeName As String)
    ' Snapshot all required list text while NO modal UserForm is active.
    ' The v10.4 and v10.5 failures occurred because a sheet object was acquired
    ' from inside UserForm_Activate (Err 91). There are no worksheet calls
    ' during initial list population in the v10.6 form.
    Dim stage As String, errNum As Long, errText As String
    On Error GoTo OpenFailed
    mMode = modeName
    Set mHostBook = ThisWorkbook
    If mHostBook Is Nothing Then Err.Raise vbObjectError + 2404, , "Робочу книгу не знайдено."
    mUiLoading = True
    stage = "Підготовка даних до відкриття форми"
    PrepareFormLists stage
    mUiLoading = False
    stage = "Відкриття форми"
    frmTCAdmin.Show vbModal
    Exit Sub
OpenFailed:
    errNum = Err.Number: errText = Err.Description
    mUiLoading = False
    MsgBox "Не вдалося відкрити форму адміністратора [" & stage & "] (" & errNum & "): " & errText, vbExclamation, "Транспортний центр"
End Sub

Private Sub PrepareFormLists(ByRef stage As String)
    ' Load ONLY the reference lists needed for the selected operation.
    ' All data are scalar Strings, never raw Worksheet/Range/COM references.
    Dim w As Worksheet, rowId As Long, lastRow As Long, code As String, caption As String
    Set mUiDrivers = New Collection
    Set mUiVehicles = New Collection
    Set mUiRoutes = New Collection
    Set mUiParkings = New Collection
    Set mUiTrips = New Collection
    mUiSummary = ""

    If mMode = "NEW" Or mMode = "DRIVER" Then
        stage = "Підготовка водіїв"
        Set w = WS("Контакти_водіїв")
        lastRow = w.Cells(w.Rows.Count, 1).End(xlUp).Row
        If lastRow > 10000 Then Err.Raise vbObjectError + 2450, , "Довідник водіїв має понад 10000 рядків."
        For rowId = FIRST_DATA To lastRow
            stage = "Водії: рядок " & CStr(rowId)
            code = UiText(w.Cells(rowId, 1).Value2)
            If Len(code) > 0 Then
                caption = UiText(w.Cells(rowId, 8).Value2)
                If Len(caption) = 0 Then caption = code
                mUiDrivers.Add caption
            End If
        Next rowId
    End If

    If mMode = "NEW" Or mMode = "DRIVER" Or mMode = "CAR" Then
        stage = "Підготовка автомобілів"
        Set w = WS("Автопарк")
        lastRow = w.Cells(w.Rows.Count, 1).End(xlUp).Row
        If lastRow > 10000 Then Err.Raise vbObjectError + 2451, , "Довідник автомобілів має понад 10000 рядків."
        For rowId = FIRST_DATA To lastRow
            stage = "Автомобілі: рядок " & CStr(rowId)
            code = UiText(w.Cells(rowId, 1).Value2)
            If Len(code) > 0 Then
                mUiVehicles.Add code & " | " & UiText(w.Cells(rowId, 2).Value2) & " " & UiText(w.Cells(rowId, 4).Value2)
            End If
        Next rowId
    End If

    If mMode = "NEW" Or mMode = "ROUTE" Then
        stage = "Підготовка маршрутів"
        Set w = WS("Маршрути_Європа")
        lastRow = w.Cells(w.Rows.Count, 1).End(xlUp).Row
        If lastRow > 10000 Then Err.Raise vbObjectError + 2452, , "Довідник маршрутів має понад 10000 рядків."
        For rowId = FIRST_DATA To lastRow
            stage = "Маршрути: рядок " & CStr(rowId)
            code = UiText(w.Cells(rowId, 1).Value2)
            If Len(code) > 0 Then
                mUiRoutes.Add code & " | " & UiText(w.Cells(rowId, 2).Value2) & " → " & UiText(w.Cells(rowId, 3).Value2)
            End If
        Next rowId
    End If

    If mMode = "COST" Or mMode = "PARKING" Then
        stage = "Підготовка паркінгів"
        Set w = WS("TC_Паркінги")
        lastRow = w.Cells(w.Rows.Count, 1).End(xlUp).Row
        If lastRow > 10000 Then Err.Raise vbObjectError + 2453, , "Довідник паркінгів має понад 10000 рядків."
        For rowId = FIRST_DATA To lastRow
            stage = "Паркінги: рядок " & CStr(rowId)
            code = UiText(w.Cells(rowId, 1).Value2)
            If Len(code) > 0 Then mUiParkings.Add code & " | " & UiText(w.Cells(rowId, 2).Value2)
        Next rowId
    End If

    If mMode = "START" Or mMode = "FINISH" Or mMode = "COST" Then
        stage = "Підготовка рейсів"
        Set w = WS("TC_Рейси")
        lastRow = w.Cells(w.Rows.Count, 1).End(xlUp).Row
        If lastRow > MAX_TRIP_ROWS Then Err.Raise vbObjectError + 2454, , "Журнал рейсів перевищує налаштований діапазон."
        For rowId = FIRST_DATA To lastRow
            stage = "Рейси: рядок " & CStr(rowId)
            code = UiText(w.Cells(rowId, 1).Value2)
            If Len(code) > 0 Then
                Select Case mMode
                    Case "START"
                        If UiText(w.Cells(rowId, 8).Value2) = "Заплановано" Then mUiTrips.Add code & " | " & UiText(w.Cells(rowId, 5).Value2) & " | " & UiText(w.Cells(rowId, 3).Value2)
                    Case "FINISH"
                        If UiText(w.Cells(rowId, 8).Value2) = "У рейсі" Then mUiTrips.Add code & " | " & UiText(w.Cells(rowId, 5).Value2) & " | " & UiText(w.Cells(rowId, 3).Value2)
                    Case "COST"
                        If UiText(w.Cells(rowId, 8).Value2) <> "Скасовано" Then mUiTrips.Add code & " | " & UiText(w.Cells(rowId, 5).Value2) & " | " & UiText(w.Cells(rowId, 3).Value2)
                End Select
            End If
        Next rowId
    End If

    If mMode = "REPORT" Then
        stage = "Підготовка фінансового огляду"
        mUiSummary = "Водіїв: " & CStr(Application.WorksheetFunction.CountA(WS("Контакти_водіїв").Range("A6:A1000"))) & _
            " | Рейсів: " & CStr(Application.WorksheetFunction.CountA(WS("TC_Рейси").Range("A6:A50000"))) & _
            " | Паркінгів: " & CStr(Application.WorksheetFunction.CountA(WS("TC_Паркінги").Range("A6:A1000")))
    End If
    stage = "Довідники підготовлено"
End Sub

Public Sub TC_FormInit(ByVal f As Object)
    ' This routine makes NO Excel worksheet calls. Values were prepared by ShowTC.
    Dim stage As String, errorNumber As Long, errorText As String, entry As Variant
    On Error GoTo Failed
    mUiLoading = True
    stage = "Підготовка елементів форми"
    f.Controls("cmdSave").Enabled = False
    f.Controls("lblTitle").Caption = "ТРАНСПОРТНИЙ ЦЕНТР | АДМІНІСТРАТОР"
    f.Controls("cmdSave").Caption = "Зберегти"
    f.Controls("cmdCancel").Caption = "Скасувати"
    stage = "Очищення списків"
    f.Controls("cboDriver").Clear: f.Controls("cboVehicle").Clear: f.Controls("cboRoute").Clear: f.Controls("cboTrip").Clear
    f.Controls("cboParking").Clear: f.Controls("cboKind").Clear: f.Controls("cboCurrency").Clear
    stage = "Наповнення потрібних списків"
    If mUiDrivers Is Nothing Or mUiVehicles Is Nothing Or mUiRoutes Is Nothing Or _
       mUiParkings Is Nothing Or mUiTrips Is Nothing Then
        Err.Raise vbObjectError + 2460, , "Дані довідників не підготовлені до відкриття форми."
    End If
    For Each entry In mUiDrivers: f.Controls("cboDriver").AddItem CStr(entry): Next entry
    For Each entry In mUiVehicles: f.Controls("cboVehicle").AddItem CStr(entry): Next entry
    For Each entry In mUiRoutes: f.Controls("cboRoute").AddItem CStr(entry): Next entry
    For Each entry In mUiParkings: f.Controls("cboParking").AddItem CStr(entry): Next entry
    For Each entry In mUiTrips: f.Controls("cboTrip").AddItem CStr(entry): Next entry
    If mMode = "COST" Then
        f.Controls("cboKind").AddItem "Пальне": f.Controls("cboKind").AddItem "Паркінг"
        f.Controls("cboKind").AddItem "Платні дороги": f.Controls("cboKind").AddItem "Ремонт": f.Controls("cboKind").AddItem "Інше"
        f.Controls("cboCurrency").AddItem "PLN": f.Controls("cboCurrency").AddItem "USD"
        f.Controls("cboCurrency").AddItem "EUR": f.Controls("cboCurrency").AddItem "UAH"
    End If
    stage = "Налаштування вікна: " & mMode
    ApplyFormMode f
    mUiLoading = False
    f.Controls("cmdSave").Enabled = True
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    mUiLoading = False
    On Error Resume Next
    f.Controls("cmdSave").Enabled = False
    f.Controls("lblInfo").Caption = "ПОМИЛКА ЗАВАНТАЖЕННЯ. Не вводьте дані."
    MsgBox "Помилка форми [" & stage & "] (" & errorNumber & "): " & errorText, vbExclamation, "Транспортний центр"
End Sub

Private Sub AddTripChoice(ByVal f As Object, ByVal ws As Worksheet, ByVal rowId As Long)
    f.Controls("cboTrip").AddItem CStr(ws.Cells(rowId, 1).Value2) & " | " & _
        CStr(ws.Cells(rowId, 5).Value2) & " | " & CStr(ws.Cells(rowId, 3).Value2)
End Sub

Private Sub EnableCombo(ByVal f As Object, ByVal name As String, ByVal label As String)
    f.Controls("lbl" & name).Visible = True
    f.Controls("cbo" & name).Visible = True
    f.Controls("lbl" & name).Caption = label
End Sub

Private Sub Field(ByVal f As Object, ByVal n As Integer, ByVal title As String, Optional ByVal defaultValue As String = "")
    f.Controls("lbl" & CStr(n)).Caption = title
    f.Controls("lbl" & CStr(n)).Visible = True
    f.Controls("txt" & CStr(n)).Visible = True
    f.Controls("txt" & CStr(n)).Text = defaultValue
End Sub

Private Sub ApplyFormMode(ByVal f As Object)
    Dim n As Integer, labelNames As Variant, k As Variant
    labelNames = Array("Driver", "Vehicle", "Route", "Trip", "Parking", "Kind", "Currency")
    For Each k In labelNames
        f.Controls("lbl" & CStr(k)).Visible = False
        f.Controls("cbo" & CStr(k)).Visible = False
    Next k
    For n = 1 To 8
        f.Controls("lbl" & CStr(n)).Visible = False
        f.Controls("txt" & CStr(n)).Visible = False
        f.Controls("txt" & CStr(n)).Text = ""
    Next n
    f.Controls("lblPreview").Caption = "Усі зміни виконує адміністратор. Перед записом створюється резервна копія."
    Select Case mMode
        Case "NEW"
            f.Controls("lblTitle").Caption = "СТВОРЕННЯ РЕЙСУ"
            EnableCombo f, "Driver", "Водій"
            EnableCombo f, "Vehicle", "Автомобіль"
            EnableCombo f, "Route", "Маршрут"
            Field f, 1, "Замовник"
            Field f, 2, "Дохід USD", "0"
            Field f, 3, "Виїзд план, дата", Format$(Date, "yyyy-mm-dd")
            Field f, 4, "Прибуття план, дата", Format$(Date, "yyyy-mm-dd")
            Field f, 5, "Км план (порожнє=довідник)"
            Field f, 6, "Ціна пального PLN/л (заміна)"
            Field f, 7, "Примітка"
            Field f, 8, "Резерв пального, %", "10"
            f.Controls("cmdSave").Caption = "Створити план"
        Case "START"
            f.Controls("lblTitle").Caption = "РОЗПОЧАТИ РЕЙС"
            EnableCombo f, "Trip", "Запланований рейс"
            Field f, 1, "Одометр на виїзді, км"
            Field f, 2, "Дата виїзду", Format$(Date, "yyyy-mm-dd")
            Field f, 3, "Примітка"
            f.Controls("cmdSave").Caption = "Розпочати"
        Case "FINISH"
            f.Controls("lblTitle").Caption = "ЗАВЕРШИТИ РЕЙС"
            EnableCombo f, "Trip", "Рейс у дорозі"
            Field f, 1, "Одометр на прибутті, км"
            Field f, 2, "Дата прибуття", Format$(Date, "yyyy-mm-dd")
            Field f, 3, "Кермування, год", "0"
            Field f, 4, "Простій, год", "0"
            Field f, 5, "Втрати часу, год", "0"
            Field f, 6, "Отримана оплата USD", "0"
            Field f, 7, "Примітка"
            f.Controls("cmdSave").Caption = "Завершити"
        Case "COST"
            f.Controls("lblTitle").Caption = "ДОДАТИ ВИТРАТУ"
            EnableCombo f, "Trip", "Рейс зі списку"
            EnableCombo f, "Kind", "Категорія"
            EnableCombo f, "Currency", "Валюта"
            EnableCombo f, "Parking", "Паркінг (за потреби)"
            Field f, 1, "Сума у валюті"
            Field f, 2, "Дата витрати", Format$(Date, "yyyy-mm-dd")
            Field f, 3, "Квитанція / коментар"
            f.Controls("cmdSave").Caption = "Зберегти витрату"
        Case "CAR"
            f.Controls("lblTitle").Caption = "КАРТКА АВТОМОБІЛЯ"
            EnableCombo f, "Vehicle", "Обрати для редагування (або створити)"
            Field f, 1, "Код автомобіля"
            Field f, 2, "Марка"
            Field f, 3, "Модель"
            Field f, 4, "Держ. номер"
            Field f, 5, "Норма л/100 км"
            Field f, 6, "Ціна пального PLN/л"
            Field f, 7, "Тип пального", "Дизель"
            Field f, 8, "Примітка"
            f.Controls("cmdSave").Caption = "Додати автомобіль"
        Case "DRIVER"
            f.Controls("lblTitle").Caption = "КАРТКА ВОДІЯ"
            EnableCombo f, "Driver", "Редагувати водія (або створити)"
            EnableCombo f, "Vehicle", "Основний автомобіль"
            Field f, 1, "Прізвище"
            Field f, 2, "Ім'я"
            Field f, 3, "Телефон"
            Field f, 4, "Електронна пошта"
            f.Controls("cmdSave").Caption = "Зареєструвати водія"
        Case "ROUTE"
            f.Controls("lblTitle").Caption = "ДОВІДНИК МАРШРУТІВ"
            EnableCombo f, "Route", "Редагувати маршрут (або створити)"
            Field f, 1, "Код маршруту"
            Field f, 2, "Звідки"
            Field f, 3, "Куди"
            Field f, 4, "Відстань, км"
            Field f, 5, "Джерело / примітка"
            f.Controls("cmdSave").Caption = "Додати маршрут"
        Case "PARKING"
            f.Controls("lblTitle").Caption = "ДОВІДНИК ПАРКІНГІВ"
            EnableCombo f, "Parking", "Редагувати паркінг (або створити)"
            Field f, 1, "Код паркінгу"
            Field f, 2, "Назва"
            Field f, 3, "Країна"
            Field f, 4, "Місто"
            Field f, 5, "Тариф PLN / пакет"
            Field f, 6, "Джерело / примітка"
            f.Controls("cmdSave").Caption = "Додати паркінг"
        Case "REPORT"
            f.Controls("lblTitle").Caption = "ПОКАЗНИКИ ТРАНСПОРТНОГО ЦЕНТРУ"
            f.Controls("cmdSave").Caption = "Експорт PDF"
            f.Controls("lblPreview").Caption = mUiSummary
    End Select
End Sub

Public Sub TC_FormSelectionChanged(ByVal f As Object, ByVal whichControl As String)
    If mUiLoading Then Exit Sub
    On Error GoTo SilentExit
    Dim i As Long, rowId As Long, display As String, carId As String
    Dim rw As Worksheet, km As Double, norma As Double, price As Double, reserve As Double
    If mMode = "NEW" Then
        If whichControl = "Driver" Then
            display = DriverIdFromDisplay(UiText(f.Controls("cboDriver").Value))
            If display <> "" Then
                rowId = MasterRow("Контакти_водіїв", display)
                carId = CStr(WS("Контакти_водіїв").Cells(rowId, 5).Value2)
                If carId <> "" Then SelectComboById f.Controls("cboVehicle"), carId
            End If
        End If
        If whichControl = "Driver" Or whichControl = "Vehicle" Or whichControl = "Route" Then
            rowId = MasterRow("Маршрути_Європа", IdPrefix(UiText(f.Controls("cboRoute").Value)))
            If rowId = 0 Then Exit Sub
            km = Val(CStr(WS("Маршрути_Європа").Cells(rowId, 4).Value2))
            rowId = MasterRow("Автопарк", IdPrefix(UiText(f.Controls("cboVehicle").Value)))
            If rowId = 0 Then Exit Sub
            norma = Val(CStr(WS("Автопарк").Cells(rowId, 6).Value2))
            price = Val(CStr(WS("Автопарк").Cells(rowId, 7).Value2))
            f.Controls("lblPreview").Caption = "Прогноз: " & Format$(km, "0") & " км; " & Format$(km * norma / 100, "0.0") & _
                " л; орієнтовно " & Format$(km * norma / 100 * price * 1.1, "0.00") & " PLN (резерв 10%)."
        End If
    End If
    If mMode = "COST" And whichControl = "Parking" Then
        rowId = MasterRow("TC_Паркінги", IdPrefix(UiText(f.Controls("cboParking").Value)))
        If rowId > 0 Then
            If IsNumeric(WS("TC_Паркінги").Cells(rowId, 5).Value2) Then f.Controls("txt1").Text = CStr(WS("TC_Паркінги").Cells(rowId, 5).Value2)
            SelectComboById f.Controls("cboKind"), "Паркінг"
            SelectComboById f.Controls("cboCurrency"), "PLN"
        End If
    End If
    Select Case mMode
        Case "CAR"
            If whichControl = "Vehicle" Then
                rowId = MasterRow("Автопарк", IdPrefix(UiText(f.Controls("cboVehicle").Value)))
                If rowId > 0 Then
                    Set rw = WS("Автопарк")
                    For i = 1 To 4: f.Controls("txt" & CStr(i)).Text = CStr(rw.Cells(rowId, i).Value2): Next i
                    f.Controls("txt5").Text = CStr(rw.Cells(rowId, 6).Value2)
                    f.Controls("txt6").Text = CStr(rw.Cells(rowId, 7).Value2)
                    f.Controls("txt7").Text = CStr(rw.Cells(rowId, 5).Value2)
                    f.Controls("txt8").Text = CStr(rw.Cells(rowId, 10).Value2)
                    f.Controls("cmdSave").Caption = "Оновити автомобіль"
                End If
            End If
        Case "DRIVER"
            If whichControl = "Driver" Then
                rowId = MasterRow("Контакти_водіїв", DriverIdFromDisplay(UiText(f.Controls("cboDriver").Value)))
                If rowId > 0 Then
                    Set rw = WS("Контакти_водіїв")
                    f.Controls("txt1").Text = CStr(rw.Cells(rowId, 2).Value2)
                    f.Controls("txt2").Text = CStr(rw.Cells(rowId, 3).Value2)
                    f.Controls("txt3").Text = CStr(rw.Cells(rowId, 4).Value2)
                    f.Controls("txt4").Text = CStr(rw.Cells(rowId, 6).Value2)
                    SelectComboById f.Controls("cboVehicle"), CStr(rw.Cells(rowId, 5).Value2)
                    f.Controls("cmdSave").Caption = "Оновити картку"
                End If
            End If
        Case "ROUTE"
            If whichControl = "Route" Then
                rowId = MasterRow("Маршрути_Європа", IdPrefix(UiText(f.Controls("cboRoute").Value)))
                If rowId > 0 Then
                    Set rw = WS("Маршрути_Європа")
                    For i = 1 To 4: f.Controls("txt" & CStr(i)).Text = CStr(rw.Cells(rowId, i).Value2): Next i
                    f.Controls("txt5").Text = CStr(rw.Cells(rowId, 13).Value2)
                    f.Controls("cmdSave").Caption = "Оновити маршрут"
                End If
            End If
        Case "PARKING"
            If whichControl = "Parking" Then
                rowId = MasterRow("TC_Паркінги", IdPrefix(UiText(f.Controls("cboParking").Value)))
                If rowId > 0 Then
                    Set rw = WS("TC_Паркінги")
                    For i = 1 To 6: f.Controls("txt" & CStr(i)).Text = CStr(rw.Cells(rowId, i).Value2): Next i
                    f.Controls("cmdSave").Caption = "Оновити тариф"
                End If
            End If
    End Select
SilentExit:
End Sub

Private Sub SelectComboById(ByVal combo As Object, ByVal id As String)
    Dim i As Long
    If id = "" Then Exit Sub
    For i = 0 To combo.ListCount - 1
        If StrComp(IdPrefix(CStr(combo.List(i))), id, vbTextCompare) = 0 Then
            combo.ListIndex = i
            Exit For
        End If
    Next i
End Sub

' UserForm button calls this function synchronously.
Public Sub TC_FormSave(ByVal f As Object)
    On Error GoTo Failed
    Select Case mMode
        Case "NEW": SaveNew f
        Case "START": SaveStart f
        Case "FINISH": SaveFinish f
        Case "COST": SaveCost f
        Case "CAR": SaveCar f
        Case "DRIVER": SaveDriver f
        Case "ROUTE": SaveRoute f
        Case "PARKING": SaveParking f
        Case "REPORT": SaveReport
        Case Else: Err.Raise 2120, , "Невідомий режим форми."
    End Select
    Unload f
    Exit Sub
Failed:
    MsgBox "Операцію не виконано [" & mMode & "] (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "Транспортний центр"
End Sub

Private Sub SaveNew(ByVal f As Object)
    Dim drvId As String, carId As String, routeId As String, w As Worksheet, car As Worksheet
    Dim r As Long, routeRow As Long, carRow As Long, id As String, km As Double
    Dim consumption As Double, unitPrice As Double, fuelReserve As Double, liters As Double, fuelCost As Double
    Dim planStart As Date, planFinish As Date, income As Double
    drvId = DriverIdFromDisplay(UiText(f.Controls("cboDriver").Text))
    If drvId = "" Then Err.Raise 2201, , "Оберіть водія зі списку."
    carId = IdPrefix(UiText(f.Controls("cboVehicle").Text))
    carRow = MasterRow("Автопарк", carId)
    If carRow = 0 Then Err.Raise 2202, , "Оберіть автомобіль зі списку."
    routeId = IdPrefix(UiText(f.Controls("cboRoute").Text))
    routeRow = MasterRow("Маршрути_Європа", routeId)
    If routeRow = 0 Then Err.Raise 2203, , "Оберіть маршрут зі списку."
    If Trim$(f.Controls("txt1").Text) = "" Then Err.Raise 2204, , "Введіть замовника."
    income = NonNeg(f.Controls("txt2").Text, "Дохід USD")
    planStart = DateOrToday(f.Controls("txt3").Text)
    planFinish = DateOrToday(f.Controls("txt4").Text)
    If planFinish < planStart Then Err.Raise 2205, , "Планове прибуття раніше виїзду."
    km = NonNeg(f.Controls("txt5").Text, "Планова відстань", True)
    If km = 0 Then km = CDbl(WS("Маршрути_Європа").Cells(routeRow, 4).Value)
    If km <= 0 Then Err.Raise 2206, , "Немає відстані маршруту."
    Set car = WS("Автопарк")
    consumption = NonNeg(CStr(car.Cells(carRow, 6).Value), "Норма л/100 км")
    unitPrice = NonNeg(f.Controls("txt6").Text, "Ціна PLN/л", True)
    If unitPrice = 0 Then unitPrice = NonNeg(CStr(car.Cells(carRow, 7).Value), "Ціна PLN/л")
    fuelReserve = NonNeg(f.Controls("txt8").Text, "Резерв %", True)
    If fuelReserve > 100 Then Err.Raise 2207, , "Резерв пального понад 100%."
    liters = Round(km * consumption / 100, 2)
    fuelCost = Round(liters * unitPrice * (1 + fuelReserve / 100), 2)
    Set w = WS("TC_Рейси")
    r = Tail(w, 1, FIRST_DATA)
    If r > MAX_TRIP_ROWS Then Err.Raise 2208, , "Журнал рейсів заповнений."
    id = "TC-" & Format$(Now, "yyyymmddhhnnss") & "-" & Format$(r, "0000")
    If MasterRow("TC_Рейси", id) > 0 Then Err.Raise 2209, , "Повтор номера рейсу."
    If MsgBox("Створити запланований рейс " & id & "?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    BackupApp "before_new_trip"
    w.Cells(r, 1).Value = id
    w.Cells(r, 2).Value = Now
    w.Cells(r, 3).Value = drvId
    w.Cells(r, 4).Value = carId
    w.Cells(r, 5).Value = routeId
    w.Cells(r, 6).Value = Trim$(f.Controls("txt1").Text)
    w.Cells(r, 7).Value = planStart
    w.Cells(r, 26).Value = planFinish
    w.Cells(r, 8).Value = "Заплановано"
    w.Cells(r, 14).Value = income
    w.Cells(r, 16).Value = km
    w.Cells(r, 17).Value = liters
    w.Cells(r, 18).Value = fuelCost
    w.Cells(r, 21).Value = Trim$(f.Controls("txt7").Text)
    w.Cells(r, 19).Formula = "=SUMIF('TC_Витрати'!$B$6:$B$50000,A" & r & ",'TC_Витрати'!$H$6:$H$50000)"
    w.Cells(r, 20).Formula = "=N" & r & "-S" & r
    EventLog "CREATE_TRIP", id
    ThisWorkbook.Save
    MsgBox "Створено рейс " & id & ". Тепер його можна розпочати зі списку.", vbInformation
End Sub

Private Sub SaveStart(ByVal f As Object)
    Dim id As String, r As Long, w As Worksheet
    id = IdPrefix(UiText(f.Controls("cboTrip").Text))
    r = TripRow(id)
    If r = 0 Then Err.Raise 2210, , "Оберіть запланований рейс зі списку."
    Set w = WS("TC_Рейси")
    If w.Cells(r, 8).Value <> "Заплановано" Then Err.Raise 2211, , "Рейс уже розпочато або завершено."
    Dim odo As Double, startDate As Date
    odo = NonNeg(f.Controls("txt1").Text, "Початковий одометр")
    startDate = DateOrToday(f.Controls("txt2").Text)
    If MsgBox("Розпочати рейс " & id & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    BackupApp "before_start"
    w.Cells(r, 9).Value = startDate
    w.Cells(r, 11).Value = odo
    w.Cells(r, 8).Value = "У рейсі"
    If Len(Trim$(f.Controls("txt3").Text)) > 0 Then w.Cells(r, 21).Value = Trim$(f.Controls("txt3").Text)
    EventLog "START_TRIP", id
    ThisWorkbook.Save
    MsgBox "Рейс " & id & " розпочато.", vbInformation
End Sub

Private Sub SaveFinish(ByVal f As Object)
    Dim id As String, r As Long, w As Worksheet, odo As Double, endDate As Date, paid As Double
    id = IdPrefix(UiText(f.Controls("cboTrip").Text))
    r = TripRow(id)
    If r = 0 Then Err.Raise 2220, , "Оберіть рейс у дорозі зі списку."
    Set w = WS("TC_Рейси")
    If w.Cells(r, 8).Value <> "У рейсі" Then Err.Raise 2221, , "Спочатку потрібно розпочати рейс."
    odo = NonNeg(f.Controls("txt1").Text, "Кінцевий одометр")
    If odo < CDbl(w.Cells(r, 11).Value) Then Err.Raise 2222, , "Кінцевий одометр менший за початковий."
    endDate = DateOrToday(f.Controls("txt2").Text)
    If endDate < CDate(w.Cells(r, 9).Value) Then Err.Raise 2223, , "Дата завершення раніше виїзду."
    paid = NonNeg(f.Controls("txt6").Text, "Оплата USD")
    If paid > CDbl(w.Cells(r, 14).Value) Then Err.Raise 2224, , "Оплата перевищує нарахування."
    Dim driving As Double, waiting As Double, lost As Double
    driving = NonNeg(f.Controls("txt3").Text, "Кермування")
    waiting = NonNeg(f.Controls("txt4").Text, "Простій")
    lost = NonNeg(f.Controls("txt5").Text, "Втрати")
    If MsgBox("Завершити рейс " & id & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    BackupApp "before_finish"
    w.Cells(r, 12).Value = odo
    w.Cells(r, 13).Value = odo - CDbl(w.Cells(r, 11).Value)
    w.Cells(r, 10).Value = endDate
    w.Cells(r, 15).Value = paid
    w.Cells(r, 8).Value = "Завершено"
    w.Cells(r, 22).Value = Now
    w.Cells(r, 23).Value = driving
    w.Cells(r, 24).Value = waiting
    w.Cells(r, 25).Value = lost
    If Len(Trim$(f.Controls("txt7").Text)) > 0 Then w.Cells(r, 21).Value = Trim$(f.Controls("txt7").Text)
    EventLog "FINISH_TRIP", id
    ThisWorkbook.Save
    MsgBox "Рейс завершено. Пробіг: " & Format$(odo - CDbl(w.Cells(r, 11).Value), "0") & " км.", vbInformation
End Sub

Private Sub SaveCost(ByVal f As Object)
    Dim id As String, tr As Long, c As Worksheet, r As Long, amount As Double, cur As String, rate As Double, usd As Double
    id = IdPrefix(UiText(f.Controls("cboTrip").Text))
    tr = TripRow(id)
    If tr = 0 Then Err.Raise 2230, , "Оберіть рейс зі списку."
    If Trim$(f.Controls("cboKind").Text) = "" Then Err.Raise 2231, , "Оберіть категорію витрат."
    cur = Trim$(f.Controls("cboCurrency").Text)
    If cur = "" Then Err.Raise 2232, , "Оберіть валюту."
    amount = NonNeg(f.Controls("txt1").Text, "Сума")
    If amount <= 0 Then Err.Raise 2233, , "Сума має бути більшою за нуль."
    Dim spendingDate As Date, parkingId As String
    spendingDate = DateOrToday(f.Controls("txt2").Text)
    parkingId = IdPrefix(UiText(f.Controls("cboParking").Text))
    If f.Controls("cboKind").Text = "Паркінг" And MasterRow("TC_Паркінги", parkingId) = 0 Then Err.Raise 2234, , "Оберіть паркінг зі списку."
    usd = 0: rate = 0
    Select Case cur
        Case "USD": usd = amount
        Case "PLN"
            If Not IsNumeric(WS("Налаштування").Range("B8").Value) Then Err.Raise 2235, , "Немає курсу USD/PLN."
            rate = CDbl(WS("Налаштування").Range("B8").Value)
            If rate <= 0 Then Err.Raise 2236, , "Курс USD/PLN не встановлено."
            usd = Round(amount / rate, 2)
        Case Else
            Err.Raise 2237, , "Для EUR/UAH потрібно вручну додати курс. Зараз доступні USD і PLN."
    End Select
    Set c = WS("TC_Витрати")
    r = Tail(c, 1, FIRST_DATA)
    If MsgBox("Додати витрату до " & id & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    BackupApp "before_expense"
    c.Cells(r, 1).Value = Now
    c.Cells(r, 2).Value = id
    c.Cells(r, 3).Value = f.Controls("cboKind").Text
    If f.Controls("cboKind").Text = "Паркінг" Then c.Cells(r, 4).Value = parkingId
    c.Cells(r, 5).Value = amount
    c.Cells(r, 6).Value = cur
    c.Cells(r, 7).Value = rate
    c.Cells(r, 8).Value = usd
    c.Cells(r, 9).Value = spendingDate
    c.Cells(r, 10).Value = Trim$(f.Controls("txt3").Text)
    c.Cells(r, 11).Value = Environ$("USERNAME")
    EventLog "ADD_EXPENSE", id
    ThisWorkbook.Save
    MsgBox "Витрата збережена. USD: " & Format$(usd, "0.00"), vbInformation
End Sub

Private Sub SaveCar(ByVal f As Object)
    Dim w As Worksheet, carId As String, r As Long, modifying As Boolean, existingId As String
    Set w = WS("Автопарк")
    carId = Trim$(f.Controls("txt1").Text)
    existingId = IdPrefix(UiText(f.Controls("cboVehicle").Value))
    modifying = Len(existingId) > 0 And MasterRow("Автопарк", existingId) > 0
    If carId = "" Or Len(carId) > 24 Then Err.Raise 2240, , "Вкажіть короткий код автомобіля."
    If modifying Then
        If StrComp(carId, existingId, vbTextCompare) <> 0 Then Err.Raise 2241, , "Код автомобіля не можна змінити, щоб не зламати старі рейси."
        r = MasterRow("Автопарк", existingId)
    Else
        If MasterRow("Автопарк", carId) > 0 Then Err.Raise 2241, , "Автомобіль уже є в довіднику."
        r = Tail(w, 1, FIRST_DATA)
    End If
    If Trim$(f.Controls("txt2").Text) = "" Or Trim$(f.Controls("txt4").Text) = "" Then Err.Raise 2242, , "Вкажіть марку й держ. номер."
    Dim norma As Double, price As Double
    norma = NonNeg(f.Controls("txt5").Text, "Норма л/100 км")
    price = NonNeg(f.Controls("txt6").Text, "Ціна PLN/л")
    If MsgBox(IIf(modifying, "Оновити", "Додати") & " автомобіль " & carId & "?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    BackupApp "before_car"
    w.Cells(r, 1).Value = carId: w.Cells(r, 2).Value = Trim$(f.Controls("txt2").Text)
    w.Cells(r, 3).Value = Trim$(f.Controls("txt3").Text): w.Cells(r, 4).Value = Trim$(f.Controls("txt4").Text)
    w.Cells(r, 5).Value = Trim$(f.Controls("txt7").Text): w.Cells(r, 6).Value = norma
    w.Cells(r, 7).Value = price: w.Cells(r, 9).Value = "Активне"
    w.Cells(r, 10).Value = Trim$(f.Controls("txt8").Text)
    EventLog IIf(modifying, "EDIT_VEHICLE", "ADD_VEHICLE"), carId
    ThisWorkbook.Save
    MsgBox "Картку автомобіля збережено.", vbInformation
End Sub

Private Sub SaveDriver(ByVal f As Object)
    Dim surname As String, given As String, carId As String, carRow As Long, plate As String
    Dim dir As Worksheet, reg As Worksheet, sh As Worksheet, k As Long, id As String, rowD As Long, rowR As Long
    surname = Trim$(f.Controls("txt1").Text): given = Trim$(f.Controls("txt2").Text)
    If surname = "" Or given = "" Then Err.Raise 2250, , "Введіть прізвище й ім'я."
    carId = IdPrefix(UiText(f.Controls("cboVehicle").Text))
    carRow = MasterRow("Автопарк", carId)
    If carRow = 0 Then Err.Raise 2251, , "Оберіть автомобіль зі списку."
    Set dir = WS("Контакти_водіїв")
    Set reg = WS("Реєстр_водіїв")
    Dim editId As String, updateRow As Long
    editId = DriverIdFromDisplay(UiText(f.Controls("cboDriver").Value))
    If editId <> "" Then
        updateRow = MasterRow("Контакти_водіїв", editId)
        If updateRow > 0 Then
            If MsgBox("Оновити дані водія " & editId & "?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
            BackupApp "before_edit_driver"
            dir.Cells(updateRow, 2).Value = surname
            dir.Cells(updateRow, 3).Value = given
            dir.Cells(updateRow, 4).Value = Trim$(f.Controls("txt3").Text)
            dir.Cells(updateRow, 5).Value = carId
            dir.Cells(updateRow, 6).Value = Trim$(f.Controls("txt4").Text)
            WS(editId).Range("B5").Value = surname & " " & given
            WS(editId).Range("E5").Value = CStr(WS("Автопарк").Cells(carRow, 4).Value)
            EventLog "EDIT_DRIVER", editId
            ThisWorkbook.Save
            MsgBox "Дані водія оновлено.", vbInformation
            Exit Sub
        End If
    End If
    rowD = Tail(dir, 1, FIRST_DATA)
    rowR = Tail(reg, 1, 13)
    If rowR > 52 Then Err.Raise 2252, , "Вичерпано 40 місць старого реєстру."
    For k = 1 To 999
        id = "Водій_" & Format$(k, "00")
        On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(id)
        On Error GoTo 0
        If sh Is Nothing Then Exit For
        Set sh = Nothing
    Next k
    If MsgBox("Створити картку " & surname & " " & given & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    BackupApp "before_driver"
    Dim template As Worksheet, savedVisibility As XlSheetVisibility
    Set template = WS("ШАБЛОН_ВОДІЯ")
    savedVisibility = template.Visible
    template.Visible = xlSheetVisible
    template.Copy Before:=WS("КІНЕЦЬ_ВОДІЇВ")
    template.Visible = savedVisibility
    Set sh = ActiveSheet
    sh.Name = id
    sh.Range("B5").Value = surname & " " & given
    plate = CStr(WS("Автопарк").Cells(carRow, 4).Value)
    sh.Range("E5").Value = plate
    reg.Cells(rowR, 1).Value = id
    dir.Cells(rowD, 1).Value = id
    dir.Cells(rowD, 2).Value = surname
    dir.Cells(rowD, 3).Value = given
    On Error Resume Next
    dir.Cells(rowD, 4).NumberFormat = "@"
    Err.Clear
    On Error GoTo 0
    dir.Cells(rowD, 4).Value = Trim$(f.Controls("txt3").Text)
    dir.Cells(rowD, 5).Value = carId
    dir.Cells(rowD, 6).Value = Trim$(f.Controls("txt4").Text)
    dir.Cells(rowD, 7).Value = "Активний"
    dir.Cells(rowD, 8).Formula = "=B" & rowD & "&"" ""&C" & rowD & "&"" [""&A" & rowD & "&""]"""
    sh.Visible = xlSheetHidden
    WS("Зміст").Activate
    EventLog "ADD_DRIVER", id
    ThisWorkbook.Save
    MsgBox "Водія додано й зареєстровано: " & id, vbInformation
End Sub

Private Sub SaveRoute(ByVal f As Object)
    Dim w As Worksheet, r As Long, routeId As String, km As Double, modifying As Boolean, previousId As String
    Set w = WS("Маршрути_Європа")
    routeId = Trim$(f.Controls("txt1").Text)
    previousId = IdPrefix(UiText(f.Controls("cboRoute").Value))
    modifying = previousId <> "" And MasterRow("Маршрути_Європа", previousId) > 0
    If routeId = "" Or Len(routeId) > 24 Then Err.Raise 2260, , "Вкажіть короткий код маршруту."
    If modifying Then
        If StrComp(previousId, routeId, vbTextCompare) <> 0 Then Err.Raise 2261, , "ID маршруту не можна змінювати — його використовують старі рейси."
        r = MasterRow("Маршрути_Європа", previousId)
    Else
        If MasterRow("Маршрути_Європа", routeId) > 0 Then Err.Raise 2261, , "Маршрут уже існує."
        r = Tail(w, 1, FIRST_DATA)
    End If
    If Trim$(f.Controls("txt2").Text) = "" Or Trim$(f.Controls("txt3").Text) = "" Then Err.Raise 2262, , "Вкажіть обидва міста."
    km = NonNeg(f.Controls("txt4").Text, "Кілометраж")
    If km <= 0 Then Err.Raise 2263, , "Кілометраж має бути більший за нуль."
    If MsgBox(IIf(modifying, "Оновити", "Додати") & " маршрут " & routeId & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    BackupApp "before_route"
    w.Cells(r, 1).Value = routeId
    w.Cells(r, 2).Value = Trim$(f.Controls("txt2").Text)
    w.Cells(r, 3).Value = Trim$(f.Controls("txt3").Text)
    w.Cells(r, 4).Value = km
    w.Cells(r, 6).Value = "Редаговано адміністратором"
    w.Cells(r, 10).Value = "УТОЧНИТИ ДЛЯ TIR"
    w.Cells(r, 13).Value = Trim$(f.Controls("txt5").Text)
    If Not modifying Then
        w.Cells(r, 7).Formula = "=COUNTIF('TC_Рейси'!$E$6:$E$50000,A" & r & ")"
        w.Cells(r, 8).Formula = "=SUMIF('TC_Рейси'!$E$6:$E$50000,A" & r & ",'TC_Рейси'!$M$6:$M$50000)"
        w.Cells(r, 9).Formula = "=COUNTIFS('TC_Рейси'!$E$6:$E$50000,A" & r & ",'TC_Рейси'!$H$6:$H$50000,""Завершено"")"
    End If
    EventLog IIf(modifying, "EDIT_ROUTE", "ADD_ROUTE"), routeId
    ThisWorkbook.Save
    MsgBox "Маршрут збережено.", vbInformation
End Sub

Private Sub SaveParking(ByVal f As Object)
    Dim w As Worksheet, r As Long, id As String, price As Double, previousId As String, modifying As Boolean
    Set w = WS("TC_Паркінги")
    id = Trim$(f.Controls("txt1").Text)
    previousId = IdPrefix(UiText(f.Controls("cboParking").Value))
    modifying = previousId <> "" And MasterRow("TC_Паркінги", previousId) > 0
    If id = "" Or Len(id) > 25 Then Err.Raise 2270, , "Вкажіть код паркінгу."
    If modifying Then
        If StrComp(previousId, id, vbTextCompare) <> 0 Then Err.Raise 2271, , "Код паркінгу не можна змінити після використання."
        r = MasterRow("TC_Паркінги", previousId)
    Else
        If MasterRow("TC_Паркінги", id) > 0 Then Err.Raise 2271, , "Паркінг уже зареєстровано."
        r = Tail(w, 1, FIRST_DATA)
    End If
    If Trim$(f.Controls("txt2").Text) = "" Then Err.Raise 2272, , "Введіть назву паркінгу."
    price = NonNeg(f.Controls("txt5").Text, "Тариф PLN/пакет")
    If MsgBox(IIf(modifying, "Оновити", "Додати") & " паркінг " & id & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    BackupApp "before_parking"
    w.Cells(r, 1).Value = id
    w.Cells(r, 2).Value = Trim$(f.Controls("txt2").Text)
    w.Cells(r, 3).Value = Trim$(f.Controls("txt3").Text)
    w.Cells(r, 4).Value = Trim$(f.Controls("txt4").Text)
    w.Cells(r, 5).Value = price
    w.Cells(r, 6).Value = Trim$(f.Controls("txt6").Text)
    w.Cells(r, 7).Value = "Активний"
    EventLog IIf(modifying, "EDIT_PARKING", "ADD_PARKING"), id
    ThisWorkbook.Save
    MsgBox "Паркінг збережено.", vbInformation
End Sub

Private Sub SaveReport()
    Dim target As String, folder As String
    folder = ThisWorkbook.Path & Application.PathSeparator & "_Звіти"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    target = folder & Application.PathSeparator & "Транспортний_звіт_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    WS("Зміст").ExportAsFixedFormat xlTypePDF, target, xlQualityStandard
    MsgBox "Звіт експортовано: " & target, vbInformation
End Sub