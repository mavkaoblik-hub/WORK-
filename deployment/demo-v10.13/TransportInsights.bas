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

' The historical TC_RefreshLive/TC_ShowLiveStatus path was retired in v10.10.
' TransportStatusFX.TC_RefreshBoth is now the SINGLE source of truth for
' current operations and archive; obsolete rebuilding code was removed.

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


' Report appendices: category structure, four-currency equivalents, vehicle
' profitability and individual completed-trip ledger. All computed locally.
' Expense values are the fixed USD equivalents recorded at the transaction date.
Private Function ReportAppendices(ByVal w As Worksheet, ByVal startRow As Long, _
    ByVal trips As Worksheet, ByVal expenses As Worksheet, _
    ByVal firstDay As Date, ByVal lastDay As Date, ByVal driverFilter As String, _
    ByVal names As Object, ByVal categoryCosts As Object, ByVal incomeUSD As Double, ByVal costUSD As Double) As Long
    Dim r As Long, i As Long, row As Long, id As String, driver As String, car As String
    Dim k As Variant, cost As Double, dayFinished As Date, dayCost As Date, state As String
    Dim vehicleIncome As Object, vehicleCost As Object, vehicleTrips As Object
    Dim tripCosts As Object, tripCar As Object, tripDriver As Object, curr As Variant
    Dim n As Long, isSelected As Boolean, cars As Worksheet, carNames As Object
    Set vehicleIncome = CreateObject("Scripting.Dictionary")
    Set vehicleCost = CreateObject("Scripting.Dictionary")
    Set vehicleTrips = CreateObject("Scripting.Dictionary")
    Set tripCosts = CreateObject("Scripting.Dictionary")
    Set tripCar = CreateObject("Scripting.Dictionary")
    Set tripDriver = CreateObject("Scripting.Dictionary")
    Set carNames = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set cars = SheetOf("Автопарк")
    On Error GoTo 0
    If Not cars Is Nothing Then
        For i = FIRST_ROW To LastAt(cars, 1)
            id = Txt(cars.Cells(i, 1).Value2)
            If id <> "" Then carNames(id) = Trim$(Txt(cars.Cells(i, 2).Value2) & " " & Txt(cars.Cells(i, 3).Value2))
        Next i
    End If
    For i = FIRST_ROW To LastAt(trips, 1)
        id = Txt(trips.Cells(i, 1).Value2)
        If id <> "" Then
            tripCar(id) = Txt(trips.Cells(i, 4).Value2)
            tripDriver(id) = Txt(trips.Cells(i, 3).Value2)
        End If
    Next i
    For i = FIRST_ROW To LastAt(expenses, 1)
        id = Txt(expenses.Cells(i, 2).Value2)
        dayCost = AsDate(expenses.Cells(i, 9).Value)
        If dayCost = 0 Then dayCost = AsDate(expenses.Cells(i, 1).Value)
        If id <> "" And tripDriver.Exists(id) Then
            driver = CStr(tripDriver(id))
            If (driverFilter = "" Or driverFilter = driver) And Within(dayCost, firstDay, lastDay) Then
                cost = Num(expenses.Cells(i, 8).Value2)
                Increment tripCosts, id, cost
                car = CStr(tripCar(id))
                If car = "" Then car = "Без автомобіля"
                Increment vehicleCost, car, cost
            End If
        End If
    Next i
    For i = FIRST_ROW To LastAt(trips, 1)
        id = Txt(trips.Cells(i, 1).Value2)
        driver = Txt(trips.Cells(i, 3).Value2)
        dayFinished = AsDate(trips.Cells(i, 10).Value)
        If id <> "" And (driverFilter = "" Or driverFilter = driver) _
            And Within(dayFinished, firstDay, lastDay) Then
            car = Txt(trips.Cells(i, 4).Value2)
            If car = "" Then car = "Без автомобіля"
            Increment vehicleIncome, car, Num(trips.Cells(i, 14).Value2)
            Increment vehicleTrips, car, 1
        End If
    Next i

    r = startRow
    StyleHeader w, "A" & r & ":L" & r, "РОЗШИФРУВАННЯ ВИТРАТ ЗА КАТЕГОРІЯМИ"
    TableHeader w, r + 1, Array("Категорія", "Сума USD", "Частка витрат, %", "Коментар", "", "", "", "", "", "", "", "")
    r = r + 2
    For Each k In categoryCosts.Keys
        w.Cells(r, 1).Value = CStr(k)
        w.Cells(r, 2).Value = At(categoryCosts, CStr(k))
        If costUSD > 0 Then w.Cells(r, 3).Value = 100 * At(categoryCosts, CStr(k)) / costUSD
        w.Cells(r, 4).Value = "Фактичні витрати за квитанціями у вибраному періоді"
        w.Rows(r).RowHeight = 24
        r = r + 1
    Next k
    If r = startRow + 2 Then w.Cells(r, 1).Value = "Витрат за період немає": r = r + 1
    r = r + 2
    StyleHeader w, "A" & r & ":L" & r, "ПІДСУМКИ В ЧОТИРЬОХ ВАЛЮТАХ  |  інформаційний перерахунок"
    TableHeader w, r + 1, Array("Валюта", "Дохід", "Витрати", "Результат", "Методика", "", "", "", "", "", "", "")
    curr = Array("PLN", "USD", "EUR", "UAH")
    For n = 0 To 3
        w.Cells(r + 2 + n, 1).Value = CStr(curr(n))
        w.Cells(r + 2 + n, 2).Value = TransportStatusFX.TC_EquivalentUSD(incomeUSD, CStr(curr(n)))
        w.Cells(r + 2 + n, 3).Value = TransportStatusFX.TC_EquivalentUSD(costUSD, CStr(curr(n)))
        w.Cells(r + 2 + n, 4).Value = TransportStatusFX.TC_EquivalentUSD(incomeUSD - costUSD, CStr(curr(n)))
        w.Cells(r + 2 + n, 5).Value = "USD — облік; інші валюти — за курсом на аркуші"
    Next n
    r = r + 8
    StyleHeader w, "A" & r & ":L" & r, "ЕФЕКТИВНІСТЬ АВТОПАРКУ"
    TableHeader w, r + 1, Array("Авто ID", "Марка та модель", "Рейсів завершено", "Дохід USD", "Витрати USD", "Результат USD", "Дохід/рейс", "", "", "", "", "")
    r = r + 2
    For Each k In tripCar.Keys
        ' Vehicle names are grouped by ID; only output each unique ID once.
        car = CStr(tripCar(k))
        If car = "" Then car = "Без автомобіля"
        If vehicleIncome.Exists(car) Or vehicleCost.Exists(car) Then
            If Not vehicleTrips.Exists("PRINTED:" & car) Then
                vehicleTrips("PRINTED:" & car) = 1
                w.Cells(r, 1).Value = car
                If carNames.Exists(car) Then w.Cells(r, 2).Value = CStr(carNames(car))
                w.Cells(r, 3).Value = At(vehicleTrips, car)
                w.Cells(r, 4).Value = At(vehicleIncome, car)
                w.Cells(r, 5).Value = At(vehicleCost, car)
                w.Cells(r, 6).Value = At(vehicleIncome, car) - At(vehicleCost, car)
                If At(vehicleTrips, car) > 0 Then w.Cells(r, 7).Value = At(vehicleIncome, car) / At(vehicleTrips, car)
                r = r + 1
            End If
        End If
    Next k
    r = r + 2
    StyleHeader w, "A" & r & ":L" & r, "ДЕТАЛІЗАЦІЯ ЗАВЕРШЕНИХ РЕЙСІВ  |  за датою прибуття"
    TableHeader w, r + 1, Array("Рейс ID", "Водій", "Авто ID", "Маршрут ID", "Виїзд", "Прибуття", "Днів", "Км факт", "Дохід USD", "Витрати USD", "Результат USD", "Примітка")
    r = r + 2: n = 0
    For i = FIRST_ROW To LastAt(trips, 1)
        id = Txt(trips.Cells(i, 1).Value2)
        driver = Txt(trips.Cells(i, 3).Value2)
        dayFinished = AsDate(trips.Cells(i, 10).Value)
        If id <> "" And (driverFilter = "" Or driverFilter = driver) _
            And Within(dayFinished, firstDay, lastDay) Then
            w.Cells(r, 1).Value = id
            w.Cells(r, 2).Value = Friendly(names, driver)
            w.Cells(r, 3).Value = Txt(trips.Cells(i, 4).Value2)
            w.Cells(r, 4).Value = Txt(trips.Cells(i, 5).Value2)
            w.Cells(r, 5).Value = trips.Cells(i, 9).Value
            w.Cells(r, 6).Value = trips.Cells(i, 10).Value
            w.Cells(r, 7).Value = DayCount(AsDate(trips.Cells(i, 9).Value), dayFinished)
            w.Cells(r, 8).Value = Num(trips.Cells(i, 13).Value2)
            w.Cells(r, 9).Value = Num(trips.Cells(i, 14).Value2)
            w.Cells(r, 10).Value = At(tripCosts, id)
            w.Cells(r, 11).Value = Num(trips.Cells(i, 14).Value2) - At(tripCosts, id)
            w.Cells(r, 12).Value = "Витрати враховані за датою квитанції"
            r = r + 1: n = n + 1
        End If
    Next i
    If n = 0 Then w.Cells(r, 1).Value = "За цей період завершених рейсів немає": r = r + 1
    On Error Resume Next
    w.Range("E" & (r - n) & ":F" & (r - 1)).NumberFormat = "dd.mm.yyyy"
    On Error GoTo 0
    r = r + 2
    With w.Range("A" & r & ":L" & (r + 1))
        .Merge
        .Value = "Методика: дохід, пробіг, кількість рейсів — за датою прибуття; витрати — за датою квитанції. Тому результат за період може включати витрати активних рейсів. Курси PLN/EUR/UAH показують еквіваленти з аркуша Курси валют."
        .WrapText = True
        .Font.Size = 10
        .Font.Italic = True
        .RowHeight = 25
    End With
    ReportAppendices = r + 1
End Function

 ' Complete invoice/expense ledger, in precisely the same period and driver scope
' as the company totals. No web calls or live FX conversions occur here.
Private Function AppendExpenseLedger(ByVal w As Worksheet, ByVal startRow As Long, _
    ByVal trips As Worksheet, ByVal expenses As Worksheet, ByVal firstDay As Date, _
    ByVal lastDay As Date, ByVal driverFilter As String, ByVal names As Object) As Long
    Dim owners As Object, i As Long, r As Long, n As Long
    Dim tripId As String, driverId As String, whenSpent As Date, total As Double
    Dim k As Variant, target As Range
    Set owners = CreateObject("Scripting.Dictionary")
    owners.CompareMode = vbTextCompare
    For i = FIRST_ROW To LastAt(trips, 1)
        tripId = Txt(trips.Cells(i, 1).Value2)
        If tripId <> "" Then owners(tripId) = Txt(trips.Cells(i, 3).Value2)
    Next i
    StyleHeader w, "A" & startRow & ":L" & startRow, _
        "КНИГА ВИТРАТ  |  усі квитанції у вибраному періоді"
    TableHeader w, startRow + 1, Array("Дата", "Рейс ID", "Водій", "Категорія", _
        "Паркінг", "Сума", "Валюта", "Курс запису", "Витрати USD", _
        "Підстава", "№ запису", "Примітка")
    r = startRow + 2
    For i = FIRST_ROW To LastAt(expenses, 1)
        tripId = Txt(expenses.Cells(i, 2).Value2)
        If tripId <> "" And owners.Exists(tripId) Then
            driverId = CStr(owners(tripId))
            If driverFilter = "" Or driverFilter = driverId Then
                whenSpent = AsDate(expenses.Cells(i, 9).Value)
                If whenSpent = 0 Then whenSpent = AsDate(expenses.Cells(i, 1).Value)
                If Within(whenSpent, firstDay, lastDay) Then
                    w.Cells(r, 1).Value = whenSpent
                    w.Cells(r, 2).Value = tripId
                    w.Cells(r, 3).Value = Friendly(names, driverId)
                    w.Cells(r, 4).Value = Txt(expenses.Cells(i, 3).Value2)
                    w.Cells(r, 5).Value = Txt(expenses.Cells(i, 4).Value2)
                    w.Cells(r, 6).Value = Num(expenses.Cells(i, 5).Value2)
                    w.Cells(r, 7).Value = Txt(expenses.Cells(i, 6).Value2)
                    w.Cells(r, 8).Value = Num(expenses.Cells(i, 7).Value2)
                    w.Cells(r, 9).Value = Num(expenses.Cells(i, 8).Value2)
                    w.Cells(r, 10).Value = "Запис із журналу"
                    w.Cells(r, 11).Value = i
                    w.Cells(r, 12).Value = Txt(expenses.Cells(i, 10).Value2)
                    total = total + Num(expenses.Cells(i, 8).Value2)
                    Set target = w.Range(w.Cells(r, 1), w.Cells(r, 12))
                    target.Interior.Color = IIf(r Mod 2 = 0, RGB(233, 243, 236), RGB(255, 255, 255))
                    w.Rows(r).RowHeight = 22
                    r = r + 1
                    n = n + 1
                End If
            End If
        End If
    Next i
    If n = 0 Then
        w.Cells(r, 1).Value = "За вибраний період квитанцій немає"
        r = r + 1
    End If
    w.Cells(r, 1).Value = "РАЗОМ ВИТРАТИ (USD)"
    w.Cells(r, 9).Value = total
    With w.Range(w.Cells(r, 1), w.Cells(r, 12))
        .Interior.Color = RGB(215, 235, 222)
        .Font.Bold = True
        .RowHeight = 29
    End With
    On Error Resume Next
    If n > 0 Then
        w.Range(w.Cells(startRow + 2, 1), w.Cells(r - 1, 1)).NumberFormat = "dd.mm.yyyy"
        w.Range(w.Cells(startRow + 2, 6), w.Cells(r, 9)).NumberFormat = "#,##0.00"
    End If
    On Error GoTo 0
    AppendExpenseLedger = r + 1
End Function

' The graph's actual source data, with the same month/driver filters as the charts.
Private Function AppendMonthlySummary(ByVal w As Worksheet, ByVal startRow As Long, _
    ByVal monthStart As Date, ByVal monthCount As Long, _
    ByVal revenue As Object, ByVal spending As Object) As Long
    Dim i As Long, r As Long, monthLabel As String, inc As Double, outflow As Double
    StyleHeader w, "A" & startRow & ":L" & startRow, "ПОМІСЯЧНА ДИНАМІКА  |  джерела даних для графіка"
    TableHeader w, startRow + 1, Array("Місяць", "Дохід USD", "Витрати USD", "Результат USD", _
        "Маржа %", "Пояснення", "", "", "", "", "", "")
    r = startRow + 2
    For i = 0 To monthCount - 1
        monthLabel = Format$(DateAdd("m", i, monthStart), "yyyy-mm")
        inc = At(revenue, monthLabel)
        outflow = At(spending, monthLabel)
        w.Cells(r, 1).Value = monthLabel
        w.Cells(r, 2).Value = inc
        w.Cells(r, 3).Value = outflow
        w.Cells(r, 4).Value = inc - outflow
        If inc > 0 Then w.Cells(r, 5).Value = (inc - outflow) * 100 / inc
        w.Cells(r, 6).Value = "Дохід за прибуттям; витрати за квитанціями"
        w.Rows(r).RowHeight = 23
        w.Range(w.Cells(r, 1), w.Cells(r, 12)).Interior.Color = _
            IIf(i Mod 2 = 0, RGB(233, 243, 236), RGB(255, 255, 255))
        r = r + 1
    Next i
    AppendMonthlySummary = r + 1
End Function

Private Function EarliestActualMonth(ByVal latest As Date, ByVal revenues As Object, _
    ByVal spending As Object) As Date
    Dim k As Variant, d As Date, first As Date
    first = DateSerial(Year(latest), Month(latest), 1)
    For Each k In revenues.Keys
        If Len(CStr(k)) = 7 Then
            d = DateSerial(CLng(Left$(CStr(k), 4)), CLng(Right$(CStr(k), 2)), 1)
            If d < first Then first = d
        End If
    Next k
    For Each k In spending.Keys
        If Len(CStr(k)) = 7 Then
            d = DateSerial(CLng(Left$(CStr(k), 4)), CLng(Right$(CStr(k), 2)), 1)
            If d < first Then first = d
        End If
    Next k
    EarliestActualMonth = first
End Function

Public Sub TC_GenerateReport(ByVal reportMode As String, ByVal startDate As Date, ByVal endDate As Date, Optional ByVal driverFilter As String = "")
    Dim w As Worksheet, trips As Worksheet, expenses As Worksheet, drvNames As Object
    Dim dIncome As Object, dPaid As Object, dCost As Object, dKm As Object, dCount As Object, dDays As Object, dActive As Object
    Dim mapTripDriver As Object, dCategory As Object, dMonthlyIncome As Object, dMonthlyExpense As Object
    Dim i As Long, r As Long, lastRow As Long, lastDriverRow As Long, monthCount As Long, monthStart As Date
    Dim id As String, driverId As String, status As String, dateFinished As Date, dateStarted As Date, whenSpent As Date, key As String
    Dim earn As Double, paid As Double, cost As Double, totalKm As Double, totalDays As Double
    Dim closed As Long, active As Long, planned As Long, idx As Long, rowDriver As Long
    Dim days As Long, resultRow As Long, endRow As Long, exportTo As String, folder As String
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
    If Len(driverFilter) > 0 Then
        labelPeriod = labelPeriod & " | Водій: " & Friendly(drvNames, driverFilter)
    Else
        labelPeriod = labelPeriod & " | Усі водії (компанія)"
    End If
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
    If Len(driverFilter) = 0 Then
        StyleHeader w, "A15:L15", "ГРАФІКИ КОМПАНІЇ  |  дохід і витрати за датами операцій"
    Else
        StyleHeader w, "A15:L15", "ГРАФІКИ ВОДІЯ  |  дохід і витрати за датами операцій"
    End If
    monthCount = DateDiff("m", DateSerial(Year(startDate), Month(startDate), 1), DateSerial(Year(endDate), Month(endDate), 1)) + 1
    If reportMode = "ALL" Then
        cardFirst = EarliestActualMonth(endDate, dMonthlyIncome, dMonthlyExpense)
        monthCount = DateDiff("m", cardFirst, DateSerial(Year(endDate), Month(endDate), 1)) + 1
    End If
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
    If Len(driverFilter) = 0 Then
        StyleHeader w, "A32:L32", "ПОКАЗНИКИ ЗА ВСІМА ВОДІЯМИ  |  порівняльна таблиця"
    Else
        StyleHeader w, "A32:L32", "ПОКАЗНИКИ ВИБРАНОГО ВОДІЯ  |  персональні результати"
    End If
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
    If Len(driverFilter) = 0 Then
        StyleHeader w, "A" & resultRow & ":L" & resultRow, "ПОРІВНЯННЯ ВОДІЇВ ЗА ДОХОДОМ ТА ВИТРАТАМИ"
    Else
        StyleHeader w, "A" & resultRow & ":L" & resultRow, "ФІНАНСОВІ ПОКАЗНИКИ ОБРАНОГО ВОДІЯ"
    End If
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
    endRow = ReportAppendices(w, resultRow + 19, trips, expenses, startDate, endDate, driverFilter, drvNames, dCategory, earn, cost)
    endRow = AppendMonthlySummary(w, endRow + 3, monthStart, monthCount, dMonthlyIncome, dMonthlyExpense)
    endRow = AppendExpenseLedger(w, endRow + 3, trips, expenses, startDate, endDate, driverFilter, drvNames)
    w.Columns("N:U").Hidden = True
    w.PageSetup.PrintArea = "$A$1:$L$" & endRow
    ' Keep table titles with following rows when Excel paginates long reports.
    On Error Resume Next
    w.PageSetup.PrintTitleRows = "$1:$3"
    On Error GoTo Failed
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
    If Len(driverFilter) = 0 Then
        exportTo = folder & Application.PathSeparator & "Transport_Report_ALL_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    Else
        exportTo = folder & Application.PathSeparator & "Transport_Report_DRIVER_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
    End If
    w.ExportAsFixedFormat Type:=xlTypePDF, Filename:=exportTo, Quality:=xlQualityStandard, IgnorePrintAreas:=False
    w.Activate
    ThisWorkbook.Save
    MsgBox "Звіт сформовано з графіками й даними за водіями:" & vbCrLf & exportTo, vbInformation, "Транспортний центр"
    Exit Sub
Failed:
    Err.Raise vbObjectError + 3131, "TransportInsights.TC_GenerateReport", "Звіт не сформовано: " & Err.Description
End Sub