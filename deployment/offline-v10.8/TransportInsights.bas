Attribute VB_Name = "TransportInsights"
Option Explicit

' Reports and live operations. This module never requests Internet access.
' Dates are taken from completed trips (arrival), active trip starts and expense transactions.
' Actual on-road mileage exists only after an administrator records an odometer update.
Private Const FIRST_ROW As Long = 6
Private Const MAX_VISIBLE_CLOSED As Long = 40

Private Function SheetOf(ByVal name As String) As Worksheet
    On Error GoTo Missing
    Set SheetOf = ThisWorkbook.Worksheets(name)
    Exit Function
Missing:
    Err.Raise vbObjectError + 3101, "TransportInsights", "Немає аркуша: " & name
End Function

Private Function LastAt(ByVal w As Worksheet, ByVal col As Long) As Long
    LastAt = w.Cells(w.Rows.Count, col).End(xlUp).Row
    If LastAt < FIRST_ROW Then LastAt = FIRST_ROW - 1
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    Txt = Trim$(CStr(v))
End Function

Private Function Num(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then Num = CDbl(v)
End Function

Private Function HasDate(ByVal v As Variant) As Boolean
    On Error GoTo NoDate
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    If Len(Txt(v)) = 0 Then Exit Function
    If IsDate(v) Then HasDate = True
NoDate:
End Function

Private Function AsDate(ByVal v As Variant) As Date
    If HasDate(v) Then AsDate = CDate(v)
End Function

Private Function DayCount(ByVal startDate As Date, ByVal endDate As Date) As Long
    If startDate = 0 Or endDate = 0 Then Exit Function
    If endDate >= startDate Then DayCount = DateDiff("d", DateValue(startDate), DateValue(endDate)) + 1
End Function

Private Sub StyleHeader(ByVal w As Worksheet, ByVal target As String, ByVal message As String)
    With w.Range(target)
        .Merge
        .Value = message
        .Interior.Color = RGB(13, 36, 31)
        .Font.Color = RGB(236, 236, 212)
        .Font.Bold = True
        .Font.Size = 15
        .RowHeight = 27
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub TableHeader(ByVal w As Worksheet, ByVal rowIndex As Long, ByVal fields As Variant)
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        w.Cells(rowIndex, i + 1).Value = CStr(fields(i))
    Next i
    With w.Range(w.Cells(rowIndex, 1), w.Cells(rowIndex, UBound(fields) + 1))
        .Interior.Color = RGB(34, 83, 70)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .WrapText = True
        .RowHeight = 31
    End With
End Sub

Private Function NameById(ByVal id As String, ByVal people As Object) As String
    If people.Exists(id) Then
        NameById = CStr(people(id))
    Else
        NameById = id
    End If
End Function

Private Function DriverNames() As Object
    Dim d As Object, w As Worksheet, r As Long, id As String, label As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set w = SheetOf("Контакти_водіїв")
    For r = FIRST_ROW To LastAt(w, 1)
        id = Txt(w.Cells(r, 1).Value2)
        If Len(id) > 0 Then
            label = Trim$(Txt(w.Cells(r, 2).Value2) & " " & Txt(w.Cells(r, 3).Value2))
            If Len(label) = 0 Then label = id
            d(id) = label
        End If
    Next r
    Set DriverNames = d
End Function

Private Function RouteNames() As Object
    Dim d As Object, w As Worksheet, r As Long, id As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set w = SheetOf("Маршрути_Європа")
    For r = FIRST_ROW To LastAt(w, 1)
        id = Txt(w.Cells(r, 1).Value2)
        If Len(id) > 0 Then d(id) = Txt(w.Cells(r, 2).Value2) & " – " & Txt(w.Cells(r, 3).Value2)
    Next r
    Set RouteNames = d
End Function

Private Function Friendly(ByVal dict As Object, ByVal key As String) As String
    If dict.Exists(key) Then Friendly = CStr(dict(key)) Else Friendly = key
End Function

Private Sub LiveColumns(ByVal w As Worksheet)
    Dim i As Long
    w.UsedRange.UnMerge
    w.Cells.Clear
    For i = 1 To 11: w.Columns(i).ColumnWidth = 16: Next i
    w.Columns("A").ColumnWidth = 25
    w.Columns("B").ColumnWidth = 23
    w.Columns("C").ColumnWidth = 34
    w.Columns("D").ColumnWidth = 17
    w.Columns("E").ColumnWidth = 17
    w.Columns("J").ColumnWidth = 21
    w.Columns("K").ColumnWidth = 31
    w.Cells.Font.Name = "Times New Roman"
    w.Cells.Font.Size = 12
    w.Range("A1:K2").Interior.Color = RGB(13, 36, 31)
    With w.Range("A1:K2")
        .Merge
        .Value = "ТРАНСПОРТНИЙ ЦЕНТР  |  ПОТОЧНИЙ СТАН"
        .Font.Color = RGB(239, 233, 204)
        .Font.Size = 20
        .Font.Bold = True
        .RowHeight = 25
        .VerticalAlignment = xlCenter
    End With
    w.Range("A3:K3").Merge
    w.Range("A3").Value = "Оновлено: " & Format$(Now, "dd.mm.yyyy HH:nn") & " | Показники тільки за записаними даними; пробіг активного рейсу потребує оновлення одометра."
    w.Range("A3").Font.Color = RGB(65, 85, 78)
    w.Range("A3").Font.Italic = True
    w.Rows(3).RowHeight = 23
End Sub

Public Sub TC_RefreshLive()
    Dim w As Worksheet, trips As Worksheet, spent As Worksheet, people As Object, routes As Object
    Dim r As Long, activeRow As Long, closedRow As Long, lastRow As Long, i As Long
    Dim planned As Long, active As Long, finished As Long, kmDone As Double, costDone As Double
    Dim id As String, status As String, driver As String, routeId As String
    Dim startAt As Date, finishAt As Date, km As Double, days As Long, expCost As Double
    Dim closedList As Collection, item As Variant, completedStart As Long
    On Error GoTo Fail
    Set w = SheetOf("Поточний стан")
    Set trips = SheetOf("TC_Рейси")
    Set spent = SheetOf("TC_Витрати")
    Set people = DriverNames()
    Set routes = RouteNames()
    Set closedList = New Collection
    LiveColumns w
    activeRow = 12
    For r = FIRST_ROW To LastAt(trips, 1)
        id = Txt(trips.Cells(r, 1).Value2)
        If Len(id) > 0 Then
            status = Txt(trips.Cells(r, 8).Value2)
            Select Case status
                Case "Заплановано": planned = planned + 1
                Case "У рейсі"
                    active = active + 1
                    startAt = AsDate(trips.Cells(r, 9).Value)
                    days = DayCount(startAt, Date)
                    km = 0
                    If IsNumeric(trips.Cells(r, 27).Value2) Then
                        km = Application.Max(0, Num(trips.Cells(r, 27).Value2) - Num(trips.Cells(r, 11).Value2))
                    End If
                    expCost = Num(trips.Cells(r, 19).Value2)
                    costDone = costDone + expCost
                    driver = Txt(trips.Cells(r, 3).Value2)
                    routeId = Txt(trips.Cells(r, 5).Value2)
                    w.Cells(activeRow, 1).Value = id
                    w.Cells(activeRow, 2).Value = Friendly(people, driver)
                    w.Cells(activeRow, 3).Value = Friendly(routes, routeId)
                    w.Cells(activeRow, 4).Value = Txt(trips.Cells(r, 4).Value2)
                    w.Cells(activeRow, 5).Value = startAt
                    w.Cells(activeRow, 6).Value = days
                    w.Cells(activeRow, 7).Value = Num(trips.Cells(r, 16).Value2)
                    If IsNumeric(trips.Cells(r, 27).Value2) Then w.Cells(activeRow, 8).Value = km Else w.Cells(activeRow, 8).Value = "не внесено"
                    w.Cells(activeRow, 9).Value = expCost
                    w.Cells(activeRow, 10).Value = Num(trips.Cells(r, 14).Value2)
                    w.Cells(activeRow, 11).Value = Txt(trips.Cells(r, 21).Value2)
                    w.Range(w.Cells(activeRow, 1), w.Cells(activeRow, 11)).Interior.Color = IIf(activeRow Mod 2 = 0, RGB(236, 245, 240), RGB(250, 252, 251))
                    activeRow = activeRow + 1
                Case "Завершено"
                    finished = finished + 1
                    kmDone = kmDone + Num(trips.Cells(r, 13).Value2)
                    costDone = costDone + Num(trips.Cells(r, 19).Value2)
                    closedList.Add r
            End Select
        End If
    Next r
    StyleHeader w, "A5:C5", "ЗАПЛАНОВАНО"
    StyleHeader w, "D5:F5", "У РЕЙСІ"
    StyleHeader w, "G5:I5", "ЗАВЕРШЕНО"
    StyleHeader w, "J5:K5", "ЗАГАЛЬНИЙ ПРОБІГ"
    w.Range("A6:C7").Merge: w.Range("A6").Value = planned
    w.Range("D6:F7").Merge: w.Range("D6").Value = active
    w.Range("G6:I7").Merge: w.Range("G6").Value = finished
    w.Range("J6:K7").Merge: w.Range("J6").Value = Format$(kmDone, "#,##0") & " км (завершені)"
    With w.Range("A6:K7")
        .Font.Bold = True: .Font.Size = 17
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = RGB(226, 238, 230)
        .RowHeight = 22
    End With
    w.Range("A9:K9").Merge
    w.Range("A9").Value = "Витрати за рейсами: " & Format$(costDone, "#,##0.00") & " USD. Активних рейсів: " & active & "."
    w.Range("A9").Font.Bold = True
    StyleHeader w, "A10:K10", "АКТИВНІ РЕЙСИ  |  адміністратор оновлює пробіг кнопкою на головній панелі"
    TableHeader w, 11, Array("Рейс ID", "Водій", "Маршрут", "Авто", "Виїзд", "Днів", "Км план", "Км пройдено", "Витрати USD", "Дохід план USD", "Примітка")
    If active = 0 Then
        w.Range("A12:K12").Merge
        w.Range("A12").Value = "Наразі активних рейсів немає."
        activeRow = 13
    End If
    completedStart = activeRow + 3
    StyleHeader w, "A" & completedStart & ":K" & completedStart, "ОСТАННІ ЗАВЕРШЕНІ РЕЙСИ  |  повна історія доступна у звіті")
    TableHeader w, completedStart + 1, Array("Рейс ID", "Водій", "Маршрут", "Авто", "Виїзд", "Прибуття", "Днів", "Км факт", "Витрати USD", "Дохід USD", "Результат USD")
    closedRow = completedStart + 2
    For i = closedList.Count To Application.Max(1, closedList.Count - MAX_VISIBLE_CLOSED + 1) Step -1
        r = CLng(closedList(i))
        driver = Txt(trips.Cells(r, 3).Value2)
        routeId = Txt(trips.Cells(r, 5).Value2)
        startAt = AsDate(trips.Cells(r, 9).Value)
        finishAt = AsDate(trips.Cells(r, 10).Value)
        w.Cells(closedRow, 1).Value = Txt(trips.Cells(r, 1).Value2)
        w.Cells(closedRow, 2).Value = Friendly(people, driver)
        w.Cells(closedRow, 3).Value = Friendly(routes, routeId)
        w.Cells(closedRow, 4).Value = Txt(trips.Cells(r, 4).Value2)
        w.Cells(closedRow, 5).Value = startAt
        w.Cells(closedRow, 6).Value = finishAt
        w.Cells(closedRow, 7).Value = DayCount(startAt, finishAt)
        w.Cells(closedRow, 8).Value = Num(trips.Cells(r, 13).Value2)
        w.Cells(closedRow, 9).Value = Num(trips.Cells(r, 19).Value2)
        w.Cells(closedRow, 10).Value = Num(trips.Cells(r, 14).Value2)
        w.Cells(closedRow, 11).Value = Num(trips.Cells(r, 14).Value2) - Num(trips.Cells(r, 19).Value2)
        w.Range(w.Cells(closedRow, 1), w.Cells(closedRow, 11)).Interior.Color = IIf(closedRow Mod 2 = 0, RGB(239, 246, 243), RGB(255, 255, 255))
        closedRow = closedRow + 1
    Next i
    w.Range("E12:E" & Application.Max(closedRow, 12)).NumberFormat = "dd.mm.yyyy"
    w.Range("E" & (completedStart + 2) & ":F" & Application.Max(closedRow, completedStart + 2)).NumberFormat = "dd.mm.yyyy"
    w.Range("H12:K" & Application.Max(closedRow, 12)).HorizontalAlignment = xlRight
    w.PageSetup.PrintArea = "$A$1:$K$" & Application.Max(closedRow, 18)
    w.PageSetup.Orientation = xlLandscape
    w.PageSetup.Zoom = False
    w.PageSetup.FitToPagesWide = 1
    w.PageSetup.FitToPagesTall = False
    Exit Sub
Fail:
    Err.Raise vbObjectError + 3110, "TransportInsights.TC_RefreshLive", "Не вдалося оновити стан рейсів: " & Err.Description
End Sub

Public Sub TC_ShowLiveStatus()
    On Error GoTo Failed
    TC_RefreshLive
    SheetOf("Поточний стан").Visible = xlSheetVisible
    SheetOf("Поточний стан").Activate
    ThisWorkbook.Save
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Поточний стан"
End Sub

Public Sub TC_GoHome()
    SheetOf("Зміст").Activate
End Sub

Private Sub Increment(ByVal dict As Object, ByVal id As String, ByVal amount As Double)
    If Not dict.Exists(id) Then dict(id) = CDbl(0)
    dict(id) = CDbl(dict(id)) + amount
End Sub

Private Function At(ByVal dict As Object, ByVal id As String) As Double
    If dict.Exists(id) Then At = CDbl(dict(id))
End Function

Private Function Within(ByVal d As Date, ByVal firstDay As Date, ByVal lastDay As Date) As Boolean
    Within = (d > 0 And DateValue(d) >= DateValue(firstDay) And DateValue(d) <= DateValue(lastDay))
End Function

Private Sub ReportCard(ByVal w As Worksheet, ByVal labelCells As String, ByVal valueCells As String, ByVal label As String, ByVal value As String)
    With w.Range(labelCells)
        .Merge: .Value = label
        .Interior.Color = RGB(24, 65, 55)
        .Font.Color = RGB(240, 240, 224)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With w.Range(valueCells)
        .Merge: .Value = value
        .Interior.Color = RGB(228, 239, 231)
        .Font.Color = RGB(18, 61, 49)
        .Font.Bold = True: .Font.Size = 18
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub AddChart(ByVal w As Worksheet, ByVal sourceRange As String, ByVal leftCell As String, ByVal rightCell As String, ByVal caption As String, ByVal chartKind As Long)
    Dim anchor As Range, endAnchor As Range, obj As ChartObject
    Set anchor = w.Range(leftCell): Set endAnchor = w.Range(rightCell)
    Set obj = w.ChartObjects.Add(anchor.Left + 2, anchor.Top + 2, endAnchor.Left + endAnchor.Width - anchor.Left - 4, endAnchor.Top + endAnchor.Height - anchor.Top - 4)
    With obj.Chart
        .ChartType = chartKind
        .SetSourceData Source:=w.Range(sourceRange)
        .PlotVisibleOnly = False
        .HasTitle = True
        .ChartTitle.Text = caption
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
End Sub

Public Sub TC_GenerateReport(ByVal reportMode As String, ByVal startDate As Date, ByVal endDate As Date, Optional ByVal driverFilter As String = "")
    Dim w As Worksheet, trips As Worksheet, expenses As Worksheet, drvNames As Object
    Dim dIncome As Object, dPaid As Object, dCost As Object, dKm As Object, dCount As Object, dDays As Object, dActive As Object
    Dim mapTripDriver As Object, dCategory As Object, dMonthlyIncome As Object, dMonthlyExpense As Object
    Dim i As Long, r As Long, lastRow As Long, lastDriverRow As Long, monthCount As Long, monthStart As Date
    Dim id As String, driverId As String, status As String, dateFinished As Date, dateStarted As Date, whenSpent As Date, key As String
    Dim earn As Double, paid As Double, cost As Double, totalKm As Double, totalDays As Double
    Dim closed As Long, active As Long, planned As Long, idx As Long, rowDriver As Long
    Dim days As Long, resultRow As Long, exportTo As String, folder As String
    Dim k As Variant, categories As Variant, cardFirst As Date, labelPeriod As String
    On Error GoTo Failed
    If endDate < startDate Then Err.Raise vbObjectError + 3130, , "Дата завершення періоду раніше його початку."
    Set drvNames = DriverNames()
    Set trips = SheetOf("TC_Рейси"): Set expenses = SheetOf("TC_Витрати")
    Set dIncome = CreateObject("Scripting.Dictionary"): Set dPaid = CreateObject("Scripting.Dictionary")
    Set dCost = CreateObject("Scripting.Dictionary"): Set dKm = CreateObject("Scripting.Dictionary")
    Set dCount = CreateObject("Scripting.Dictionary"): Set dDays = CreateObject("Scripting.Dictionary")
    Set dActive = CreateObject("Scripting.Dictionary"): Set mapTripDriver = CreateObject("Scripting.Dictionary")
    Set dCategory = CreateObject("Scripting.Dictionary")
    Set dMonthlyIncome = CreateObject("Scripting.Dictionary"): Set dMonthlyExpense = CreateObject("Scripting.Dictionary")
    mapTripDriver.CompareMode = vbTextCompare
    For r = FIRST_ROW To LastAt(trips, 1)
        id = Txt(trips.Cells(r, 1).Value2)
        driverId = Txt(trips.Cells(r, 3).Value2)
        If Len(id) > 0 Then
            mapTripDriver(id) = driverId
            If Len(driverId) > 0 And Not drvNames.Exists(driverId) Then drvNames(driverId) = driverId
            If Len(driverFilter) = 0 Or driverFilter = driverId Then
                dateStarted = AsDate(trips.Cells(r, 9).Value)
                dateFinished = AsDate(trips.Cells(r, 10).Value)
                status = Txt(trips.Cells(r, 8).Value2)
                If dateFinished > 0 And DateValue(dateFinished) <= DateValue(endDate) Then
                    If Within(dateFinished, startDate, endDate) Then
                        closed = closed + 1
                        earn = earn + Num(trips.Cells(r, 14).Value2)
                        paid = paid + Num(trips.Cells(r, 15).Value2)
                        totalKm = totalKm + Num(trips.Cells(r, 13).Value2)
                        If dateStarted = 0 Then dateStarted = AsDate(trips.Cells(r, 7).Value)
                        days = DayCount(dateStarted, dateFinished): totalDays = totalDays + days
                        Increment dIncome, driverId, Num(trips.Cells(r, 14).Value2)
                        Increment dPaid, driverId, Num(trips.Cells(r, 15).Value2)
                        Increment dKm, driverId, Num(trips.Cells(r, 13).Value2)
                        Increment dDays, driverId, CDbl(days)
                        Increment dCount, driverId, 1
                        key = Format$(dateFinished, "yyyy-mm")
                        Increment dMonthlyIncome, key, Num(trips.Cells(r, 14).Value2)
                    End If
                ElseIf dateStarted > 0 And DateValue(dateStarted) <= DateValue(endDate) And status <> "Скасовано" Then
                    active = active + 1
                    Increment dActive, driverId, 1
                ElseIf status <> "Скасовано" And DateValue(AsDate(trips.Cells(r, 2).Value)) <= DateValue(endDate) Then
                    planned = planned + 1
                End If
            End If
        End If
    Next r
    For r = FIRST_ROW To LastAt(expenses, 1)
        id = Txt(expenses.Cells(r, 2).Value2)
        whenSpent = AsDate(expenses.Cells(r, 9).Value)
        If whenSpent = 0 Then whenSpent = AsDate(expenses.Cells(r, 1).Value)
        If Len(id) > 0 And mapTripDriver.Exists(id) And Within(whenSpent, startDate, endDate) Then
            driverId = CStr(mapTripDriver(id))
            If Len(driverFilter) = 0 Or driverFilter = driverId Then
                cost = cost + Num(expenses.Cells(r, 8).Value2)
                Increment dCost, driverId, Num(expenses.Cells(r, 8).Value2)
                key = Format$(whenSpent, "yyyy-mm")
                Increment dMonthlyExpense, key, Num(expenses.Cells(r, 8).Value2)
                key = Txt(expenses.Cells(r, 3).Value2)
                If Len(key) = 0 Then key = "Інші"
                Increment dCategory, key, Num(expenses.Cells(r, 8).Value2)
            End If
        End If
    Next r
    On Error Resume Next
    Set w = ThisWorkbook.Worksheets("TC_Звіт")
    On Error GoTo Failed
    If w Is Nothing Then
        Set w = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        w.Name = "TC_Звіт"
    End If
    w.Visible = xlSheetVisible
    For i = w.ChartObjects.Count To 1 Step -1: w.ChartObjects(i).Delete: Next i
    w.UsedRange.UnMerge
    w.Cells.Clear
    w.Cells.Font.Name = "Times New Roman": w.Cells.Font.Size = 11
    w.Columns("A:L").ColumnWidth = 12
    w.Columns("A").ColumnWidth = 19: w.Columns("B").ColumnWidth = 19
    w.Columns("C").ColumnWidth = 17: w.Columns("D").ColumnWidth = 14
    w.Columns("E").ColumnWidth = 15: w.Columns("F").ColumnWidth = 14
    w.Columns("N:U").ColumnWidth = 16
    w.Rows("1:76").RowHeight = 20
    StyleHeader w, "A1:L2", "ТРАНСПОРТНИЙ ЦЕНТР  |  ЗВІТ ПРО РЕЙСИ ТА ФІНАНСИ"
    If reportMode = "ALL" Then
        labelPeriod = "За весь період станом на " & Format$(endDate, "dd.mm.yyyy")
    Else
        labelPeriod = "За період з " & Format$(startDate, "dd.mm.yyyy") & " по " & Format$(endDate, "dd.mm.yyyy")
    End If
    If Len(driverFilter) > 0 Then labelPeriod = labelPeriod & " | Водій: " & Friendly(drvNames, driverFilter)
    If reportMode <> "ALL" And DateDiff("m", startDate, endDate) >= 12 Then labelPeriod = labelPeriod & " | графік за останні 12 місяців"
    w.Range("A3:L3").Merge: w.Range("A3").Value = labelPeriod
    w.Range("A3").Font.Italic = True: w.Rows(3).RowHeight = 26
    ReportCard w, "A5:C5", "A6:C8", "ДОХІД НАРАХОВАНО, USD", Format$(earn, "#,##0.00")
    ReportCard w, "D5:F5", "D6:F8", "ВИТРАТИ, USD", Format$(cost, "#,##0.00")
    ReportCard w, "G5:I5", "G6:I8", "РЕЗУЛЬТАТ, USD", Format$(earn - cost, "#,##0.00")
    ReportCard w, "J5:L5", "J6:L8", "ПРОЙДЕНО, КМ", Format$(totalKm, "#,##0")
    ReportCard w, "A10:C10", "A11:C12", "ЗАВЕРШЕНО", CStr(closed)
    ReportCard w, "D10:F10", "D11:F12", "У РЕЙСІ НА ДАТУ", CStr(active)
    ReportCard w, "G10:I10", "G11:I12", "ОТРИМАНО, USD", Format$(paid, "#,##0.00")
    ReportCard w, "J10:L10", "J11:L12", "ЛЮДИНО-ДНІ РЕЙСІВ", Format$(totalDays, "#,##0")
    StyleHeader w, "A15:L15", "ГРАФІКИ КОМПАНІЇ  |  фактичні завершення та витрати за датами операцій"
    monthCount = DateDiff("m", DateSerial(Year(startDate), Month(startDate), 1), DateSerial(Year(endDate), Month(endDate), 1)) + 1
    If reportMode = "ALL" Then monthCount = 12
    If monthCount < 1 Then monthCount = 1
    If monthCount > 12 Then monthCount = 12
    monthStart = DateSerial(Year(endDate), Month(endDate) - monthCount + 1, 1)
    w.Cells(3, 14).Value = "Місяць": w.Cells(3, 15).Value = "Дохід USD": w.Cells(3, 16).Value = "Витрати USD"
    For i = 0 To monthCount - 1
        cardFirst = DateAdd("m", i, monthStart)
        key = Format$(cardFirst, "yyyy-mm")
        w.Cells(i + 4, 14).Value = key
        w.Cells(i + 4, 15).Value = At(dMonthlyIncome, key)
        w.Cells(i + 4, 16).Value = At(dMonthlyExpense, key)
    Next i
    w.Cells(3, 18).Value = "Категорія": w.Cells(3, 19).Value = "Витрати USD"
    idx = 0
    For Each k In dCategory.Keys
        idx = idx + 1
        w.Cells(idx + 3, 18).Value = CStr(k)
        w.Cells(idx + 3, 19).Value = At(dCategory, CStr(k))
    Next k
    If idx = 0 Then
        idx = 1: w.Cells(4, 18).Value = "Немає витрат": w.Cells(4, 19).Value = 0
    End If
    AddChart w, "N3:P" & (monthCount + 3), "A16", "F29", "Нараховано vs витрати (USD)", xlColumnClustered
    AddChart w, "R3:S" & (idx + 3), "G16", "L29", "Структура витрат за категоріями", xlColumnClustered
    StyleHeader w, "A32:L32", "ПОКАЗНИКИ ЗА ВОДІЯМИ  |  доходи — за завершенням рейсу, витрати — за датою витрати"
    TableHeader w, 33, Array("Водій ID", "ПІБ", "Завершено", "У рейсі", "Км факт", "Днів", "Дохід USD", "Отримано USD", "Витрати USD", "Результат USD", "Витрати/км", "Дохід/км")
    rowDriver = 34
    For Each k In drvNames.Keys
        id = CStr(k)
        If Len(driverFilter) = 0 Or driverFilter = id Then
            w.Cells(rowDriver, 1).Value = id
            w.Cells(rowDriver, 2).Value = Friendly(drvNames, id)
            w.Cells(rowDriver, 3).Value = At(dCount, id)
            w.Cells(rowDriver, 4).Value = At(dActive, id)
            w.Cells(rowDriver, 5).Value = At(dKm, id)
            w.Cells(rowDriver, 6).Value = At(dDays, id)
            w.Cells(rowDriver, 7).Value = At(dIncome, id)
            w.Cells(rowDriver, 8).Value = At(dPaid, id)
            w.Cells(rowDriver, 9).Value = At(dCost, id)
            w.Cells(rowDriver, 10).Value = At(dIncome, id) - At(dCost, id)
            If At(dKm, id) > 0 Then
                w.Cells(rowDriver, 11).Value = At(dCost, id) / At(dKm, id)
                w.Cells(rowDriver, 12).Value = At(dIncome, id) / At(dKm, id)
            End If
            w.Range(w.Cells(rowDriver, 1), w.Cells(rowDriver, 12)).Interior.Color = IIf(rowDriver Mod 2 = 0, RGB(234, 244, 238), RGB(255, 255, 255))
            rowDriver = rowDriver + 1
        End If
    Next k
    resultRow = rowDriver + 2
    StyleHeader w, "A" & resultRow & ":L" & resultRow, "ПОРІВНЯННЯ ВОДІЇВ ЗА ДОХОДОМ ТА ВИТРАТАМИ"
    w.Cells(24, 14).Value = "Водій": w.Cells(24, 15).Value = "Дохід USD": w.Cells(24, 16).Value = "Витрати USD"
    idx = 0
    For Each k In drvNames.Keys
        id = CStr(k)
        If Len(driverFilter) = 0 Or driverFilter = id Then
            idx = idx + 1
            w.Cells(idx + 24, 14).Value = Friendly(drvNames, id)
            w.Cells(idx + 24, 15).Value = At(dIncome, id)
            w.Cells(idx + 24, 16).Value = At(dCost, id)
        End If
    Next k
    If idx = 0 Then idx = 1: w.Cells(25, 14).Value = "Немає даних": w.Cells(25, 15).Value = 0: w.Cells(25, 16).Value = 0
    AddChart w, "N24:P" & (24 + idx), "A" & (resultRow + 1), "L" & (resultRow + 16), "Водії: дохід та витрати USD", xlBarClustered
    w.Range("A" & (resultRow + 18) & ":L" & (resultRow + 18)).Merge
    w.Cells(resultRow + 18, 1).Value = "Методика: дохід і пробіг — за датою завершення; витрати — за датою квитанції. Активні рейси — станом на останній день періоду. Пробіг у дорозі не вигадується."
    w.Cells(resultRow + 18, 1).Font.Size = 10
    w.Cells(resultRow + 18, 1).WrapText = True
    w.Rows(resultRow + 18).RowHeight = 33
    w.Columns("N:U").Hidden = True
    w.PageSetup.PrintArea = "$A$1:$L$" & (resultRow + 18)
    With w.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(0.8)
        .RightMargin = Application.CentimetersToPoints(0.8)
        .TopMargin = Application.CentimetersToPoints(0.8)
        .BottomMargin = Application.CentimetersToPoints(0.8)
        .CenterFooter = "Сторінка &P з &N"
    End With
    folder = ThisWorkbook.Path & Application.PathSeparator & "_Звіти"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    exportTo = folder & Application.PathSeparator & "Transport_Report_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    w.ExportAsFixedFormat Type:=xlTypePDF, Filename:=exportTo, Quality:=xlQualityStandard, IgnorePrintAreas:=False
    w.Activate
    ThisWorkbook.Save
    MsgBox "Звіт сформовано з графіками й даними за водіями:" & vbCrLf & exportTo, vbInformation, "Транспортний центр"
    Exit Sub
Failed:
    Err.Raise vbObjectError + 3131, "TransportInsights.TC_GenerateReport", "Звіт не сформовано: " & Err.Description
End Sub