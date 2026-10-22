Attribute VB_Name = "TransportStatusFX"
Option Explicit

' Transport Center — independent operational screens and manual FX management.
' No Workbook_Open handlers, timers, recalculation hooks or background HTTP.
' Bookkeeping base: USD. Other currencies are reference equivalents at rates shown in "Курси валют".
Private Const ROW0 As Long = 6

Private Function W(ByVal sheetName As String) As Worksheet
    Set W = ThisWorkbook.Worksheets(sheetName)
End Function

Private Function S(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    S = Trim$(CStr(v))
End Function

Private Function N(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then N = CDbl(v)
End Function

Private Function Last(ByVal w As Worksheet, ByVal c As Long) As Long
    Last = w.Cells(w.Rows.Count, c).End(xlUp).Row
    If Last < ROW0 Then Last = ROW0 - 1
End Function

Private Function TripDays(ByVal departure As Variant, ByVal arrival As Variant, ByVal ongoing As Boolean) As Long
    On Error GoTo Invalid
    If Not IsDate(departure) Then Exit Function
    If ongoing Then
        If DateValue(departure) > Date Then Exit Function
        TripDays = DateDiff("d", DateValue(departure), Date) + 1
    ElseIf IsDate(arrival) Then
        If DateValue(arrival) < DateValue(departure) Then Exit Function
        TripDays = DateDiff("d", DateValue(departure), DateValue(arrival)) + 1
    End If
Invalid:
End Function

Private Function Drivers() As Object
    Dim d As Object, sheet As Worksheet, r As Long, id As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set sheet = W("Контакти_водіїв")
    For r = ROW0 To Last(sheet, 1)
        id = S(sheet.Cells(r, 1).Value2)
        If id <> "" Then d(id) = Trim$(S(sheet.Cells(r, 2).Value2) & " " & S(sheet.Cells(r, 3).Value2))
    Next r
    Set Drivers = d
End Function

Private Function Routes() As Object
    Dim d As Object, sheet As Worksheet, r As Long, id As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set sheet = W("Маршрути_Європа")
    For r = ROW0 To Last(sheet, 1)
        id = S(sheet.Cells(r, 1).Value2)
        If id <> "" Then d(id) = S(sheet.Cells(r, 2).Value2) & " → " & S(sheet.Cells(r, 3).Value2)
    Next r
    Set Routes = d
End Function

Private Function Display(ByVal catalog As Object, ByVal key As String) As String
    If catalog.Exists(key) Then
        Display = S(catalog(key))
        If Display = "" Then Display = key
    Else
        Display = key
    End If
End Function

Private Sub AddMoney(ByVal d As Object, ByVal key As String, ByVal amount As Double)
    If key = "" Then Exit Sub
    If Not d.Exists(key) Then d.Add key, CDbl(0)
    d(key) = CDbl(d(key)) + amount
End Sub

Private Function Money(ByVal d As Object, ByVal key As String) As Double
    If d.Exists(key) Then Money = CDbl(d(key))
End Function

Private Sub Header(ByVal sheet As Worksheet, ByVal kind As String, ByVal note As String)
    Dim i As Long
    ' Chart objects are presentation-only and must never block ledger rendering.
    On Error Resume Next
    For i = sheet.ChartObjects.Count To 1 Step -1
        sheet.ChartObjects(i).Delete
    Next i
    Err.Clear
    On Error GoTo 0
    ' Clear only the rendered report, not Excel's million-row sheet.
    Dim lastContentRow As Long, reportArea As Range
    lastContentRow = Application.Max(120, _
        sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row, _
        sheet.Cells(sheet.Rows.Count, 12).End(xlUp).Row)
    Set reportArea = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastContentRow + 40, 12))
    reportArea.UnMerge
    reportArea.Clear
    sheet.Range("N1").ClearContents
    With reportArea.Font
        .Name = "Calibri"
        .Size = 11
        .Color = RGB(20, 48, 40)
    End With
    sheet.Columns("A").ColumnWidth = 15
    sheet.Columns("B").ColumnWidth = 19
    sheet.Columns("C").ColumnWidth = 23
    sheet.Columns("D").ColumnWidth = 12
    sheet.Columns("E").ColumnWidth = 13
    sheet.Columns("F").ColumnWidth = 12
    sheet.Columns("G").ColumnWidth = 14
    sheet.Columns("H").ColumnWidth = 14
    sheet.Columns("I").ColumnWidth = 15
    sheet.Columns("J").ColumnWidth = 15
    sheet.Columns("K").ColumnWidth = 13
    sheet.Columns("L").ColumnWidth = 29
    sheet.Rows(4).RowHeight = 49
    sheet.Rows(6).RowHeight = 7
    sheet.Rows(9).RowHeight = 12
    sheet.Rows(11).RowHeight = 7
    With sheet.Range("A1:L2")
        .Merge
        .Value = "ТРАНСПОРТНИЙ ЦЕНТР  /  " & kind
        .Interior.Color = RGB(12, 36, 31)
        .Font.Color = RGB(244, 234, 204)
        .Font.Size = 19
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .RowHeight = 26
    End With
    With sheet.Range("A3:L3")
        .Merge
        .Value = "Станом на " & Format$(Now, "dd.mm.yyyy HH:nn") & "  |  " & note
        .Font.Size = 11
        .Font.Italic = True
        .RowHeight = 23
    End With
    With sheet.Range("A5:L5")
        .Interior.Color = RGB(39, 91, 75)
        .Font.Color = vbWhite
        .Font.Bold = True
        .RowHeight = 24
    End With
End Sub

Private Sub Bar(ByVal sheet As Worksheet, ByVal r As Long, ByVal title As String)
    With sheet.Range(sheet.Cells(r, 1), sheet.Cells(r, 12))
        .Merge
        .Value = title
        .Interior.Color = RGB(23, 71, 59)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .RowHeight = 27
    End With
End Sub

Private Sub TableLabels(ByVal sheet As Worksheet, ByVal r As Long, ByVal labels As Variant)
    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        sheet.Cells(r, i + 1).Value = CStr(labels(i))
    Next i
    With sheet.Range(sheet.Cells(r, 1), sheet.Cells(r, UBound(labels) + 1))
        .Interior.Color = RGB(52, 107, 89)
        .Font.Color = vbWhite
        .Font.Bold = True
        .WrapText = True
        .Font.Size = 11
        .VerticalAlignment = xlCenter
        .RowHeight = 44
    End With
End Sub

Private Sub Stripe(ByVal sheet As Worksheet, ByVal r As Long)
    With sheet.Range(sheet.Cells(r, 1), sheet.Cells(r, 12))
        If r Mod 2 = 0 Then .Interior.Color = RGB(234, 243, 238) Else .Interior.Color = RGB(250, 251, 250)
        .RowHeight = 29
    End With
End Sub

' Four read-only KPI cards; separate label and number rows prevent clipped text.
Private Sub KpiCard(ByVal sheet As Worksheet, ByVal startColumn As Long, ByVal caption As String, ByVal figure As String)
    Dim label As Range, value As Range
    Set label = sheet.Range(sheet.Cells(7, startColumn), sheet.Cells(7, startColumn + 2))
    Set value = sheet.Range(sheet.Cells(8, startColumn), sheet.Cells(8, startColumn + 2))
    label.Merge: value.Merge
    With label
        .Value = caption
        .Interior.Color = RGB(37, 85, 70)
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 24
    End With
    With value
        .Value = figure
        .Interior.Color = RGB(231, 241, 235)
        .Font.Name = "Calibri"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(13, 53, 40)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 39
    End With
End Sub


' Executive per-driver summary, dynamically sized and separate from detailed trip rows.
Private Sub DriverOverview(ByVal sheet As Worksheet, ByVal people As Object, ByVal firstRow As Long, _
    ByVal completedBy As Object, ByVal activeBy As Object, ByVal kmBy As Object, ByVal daysBy As Object, _
    ByVal incomeBy As Object, ByVal expensesBy As Object, ByVal finishedOnly As Boolean)
    Dim key As Variant, r As Long, id As String, doneCount As Double, earned As Double
    Dim spent As Double, traveled As Double
    TableLabels sheet, firstRow - 1, Array("Водій ID", "Водій", "Завершено", "Активні", "Км пройдено", "Середнє днів", "Дохід USD", "Витрати USD", "Результат USD", "Дохід/рейс", "Маржа %", "Призначення")
    r = firstRow
    For Each key In people.Keys
        id = CStr(key)
        doneCount = Money(completedBy, id)
        earned = Money(incomeBy, id)
        spent = Money(expensesBy, id)
        traveled = Money(kmBy, id)
        sheet.Cells(r, 1).Value = id
        sheet.Cells(r, 2).Value = Display(people, id)
        sheet.Cells(r, 3).Value = doneCount
        If finishedOnly Then
            sheet.Cells(r, 4).Value = "—"
        Else
            sheet.Cells(r, 4).Value = Money(activeBy, id)
        End If
        sheet.Cells(r, 5).Value = traveled
        If doneCount > 0 Then sheet.Cells(r, 6).Value = Money(daysBy, id) / doneCount
        sheet.Cells(r, 7).Value = earned
        sheet.Cells(r, 8).Value = spent
        sheet.Cells(r, 9).Value = earned - spent
        If doneCount > 0 Then sheet.Cells(r, 10).Value = earned / doneCount
        If earned > 0 Then sheet.Cells(r, 11).Value = (earned - spent) * 100 / earned
        sheet.Cells(r, 12).Value = IIf(finishedOnly, "Виконані рейси", "Усі рейси")
        Stripe sheet, r
        r = r + 1
    Next key
    If r = firstRow Then sheet.Cells(r, 1).Value = "Водіїв поки немає"
    On Error Resume Next
    sheet.Range(sheet.Cells(firstRow, 5), sheet.Cells(Application.Max(firstRow, r - 1), 5)).NumberFormat = "#,##0"
    sheet.Range(sheet.Cells(firstRow, 6), sheet.Cells(Application.Max(firstRow, r - 1), 6)).NumberFormat = "0.0"
    sheet.Range(sheet.Cells(firstRow, 7), sheet.Cells(Application.Max(firstRow, r - 1), 10)).NumberFormat = "#,##0.00"
    sheet.Range(sheet.Cells(firstRow, 11), sheet.Cells(Application.Max(firstRow, r - 1), 11)).NumberFormat = "0.0"
    On Error GoTo 0
End Sub

' A missing/unreadable REFERENCE exchange rate must NEVER invalidate USD ledger.
' Excel LTSC previously raised error 91 while rendering the archive FX footer.
Private Function Rate(ByVal iso As String) As Double
    Dim fxSheet As Worksheet, fxRow As Long, q As Variant
    If UCase$(Trim$(iso)) = "PLN" Then
        Rate = 1#
        Exit Function
    End If
    Select Case UCase$(Trim$(iso))
        Case "USD": fxRow = 7
        Case "EUR": fxRow = 8
        Case "UAH": fxRow = 9
        Case Else: Exit Function
    End Select
    On Error GoTo RateUnavailable
    Set fxSheet = ThisWorkbook.Worksheets("Курси валют")
    If fxSheet Is Nothing Then GoTo RateUnavailable
    q = fxSheet.Cells(fxRow, 2).Value2
    If IsError(q) Or IsEmpty(q) Or IsNull(q) Then Exit Function
    If Not IsNumeric(q) Then Exit Function
    If CDbl(q) > 0# Then Rate = CDbl(q)
    Exit Function
RateUnavailable:
    ' Missing rate appears as "немає курсу"; real revenue/expense sums survive.
    Debug.Print "TC_FX_RATE_UNAVAILABLE: " & iso & "; " & Err.Number & "; " & Err.Description
    Rate = 0#
    Err.Clear
End Function

' USD is the accounting ledger currency; PLN/EUR/UAH are display equivalents.
' One PUBLIC entry point is used by BOTH reporting modules and the financial footer.
' Avoid indirect calls through the former local "Equivalent" helper: some LTSC
' installations stopped compilation at FinancialFooter with "Sub or Function not defined".
Public Function TC_EquivalentUSD(ByVal usdAmount As Double, ByVal code As String) As String
    Dim usdRate As Double, targetRate As Double, targetISO As String
    targetISO = UCase$(Trim$(code))
    If targetISO = "USD" Then
        TC_EquivalentUSD = Format$(usdAmount, "#,##0.00")
        Exit Function
    End If
    usdRate = Rate("USD")
    targetRate = Rate(targetISO)
    If usdRate <= 0 Or targetRate <= 0 Then
        TC_EquivalentUSD = "немає курсу"
    Else
        TC_EquivalentUSD = Format$(usdAmount * usdRate / targetRate, "#,##0.00")
    End If
End Function

Private Function DisplayCategory(ByVal cache As Object, ByVal trip As String) As String
    If cache.Exists(trip) Then
        DisplayCategory = CStr(cache(trip))
    Else
        DisplayCategory = "—"
    End If
End Function

Private Function CategorySummary(ByVal d As Object, ByVal trip As String) As String
    Dim v As Variant, result As String, k As String, amount As Double
    For Each v In d.Keys
        k = CStr(v)
        If Left$(k, Len(trip) + 1) = trip & "|" Then
            amount = Money(d, k)
            If amount <> 0 Then
                If Len(result) > 0 Then result = result & "; "
                result = result & Mid$(k, Len(trip) + 2) & " " & Format$(amount, "0.00") & " USD"
            End If
        End If
    Next v
    CategorySummary = result
    If result = "" Then CategorySummary = "—"
End Function

Private Function CategoryTotalsTable(ByVal sheet As Worksheet, ByVal startRow As Long, ByVal caption As String, ByVal totals As Object) As Long
    ' Stable company-wide category summary; never count the same expense twice.
    ' Each transaction contributes its fixed USD ledger equivalent exactly once.
    Dim key As Variant, r As Long
    Bar sheet, startRow, caption
    TableLabels sheet, startRow + 1, Array("КАТЕГОРІЯ", "ВИТРАТИ USD", "Пояснення", "", "", "", "", "", "", "", "", "")
    r = startRow + 2
    For Each key In totals.Keys
        sheet.Cells(r, 1).Value = CStr(key)
        sheet.Cells(r, 2).Value = Money(totals, CStr(key))
        sheet.Cells(r, 3).Value = "Сума фактичних витрат цієї категорії"
        Stripe sheet, r
        r = r + 1
    Next key
    If r = startRow + 2 Then
        sheet.Cells(r, 1).Value = "Витрат немає"
        sheet.Cells(r, 2).Value = 0
        Stripe sheet, r
        r = r + 1
    End If
    ' Formatting is cosmetic: do not lose verified ledger totals if Excel LTSC refuses the format.
    On Error Resume Next
    sheet.Range(sheet.Cells(startRow + 2, 2), sheet.Cells(r - 1, 2)).NumberFormat = "#,##0.00"
    On Error GoTo 0
    CategoryTotalsTable = r
End Function

Private Sub FinancialFooter(ByVal sheet As Worksheet, ByVal line As Long, ByVal caption As String, ByVal revenueUSD As Double, ByVal costUSD As Double)
    Dim currencies As Variant, i As Long, fx As String
    Dim footerStep As String, originalNo As Long, originalText As String
    If sheet Is Nothing Then Err.Raise vbObjectError + 3370, "FinancialFooter", "Не передано аркуш валютного підсумку"
    On Error GoTo FooterFailed
    footerStep = "header"
    Bar sheet, line, caption
    line = line + 1
    footerStep = "currency table headers"
    TableLabels sheet, line, Array("Валюта", "ДОХІД", "ВИТРАТИ", "РЕЗУЛЬТАТ", "Перерахунок", "", "", "", "", "", "", "")
    currencies = Array("PLN", "USD", "EUR", "UAH")
    For i = 0 To 3
        footerStep = "currency " & CStr(i)
        fx = CStr(currencies(i))
        sheet.Cells(line + 1 + i, 1).Value = fx
        footerStep = fx & ": revenue"
        sheet.Cells(line + 1 + i, 2).Value = TC_EquivalentUSD(revenueUSD, fx)
        footerStep = fx & ": expense"
        sheet.Cells(line + 1 + i, 3).Value = TC_EquivalentUSD(costUSD, fx)
        footerStep = fx & ": result"
        sheet.Cells(line + 1 + i, 4).Value = TC_EquivalentUSD(revenueUSD - costUSD, fx)
        If fx = "USD" Then
            sheet.Cells(line + 1 + i, 5).Value = "Облікова валюта"
        Else
            sheet.Cells(line + 1 + i, 5).Value = "за курсом аркуша Курси валют"
        End If
        Stripe sheet, line + i + 1
        sheet.Cells(line + 1 + i, 1).Font.Bold = True
    Next i
    footerStep = "explanatory note"
    With sheet.Range(sheet.Cells(line + 6, 1), sheet.Cells(line + 6, 12))
        .Merge
        .Value = "Дохід: нарахування завершених рейсів; витрати: фактичні записи. Курс показує еквівалент, а не додатковий прибуток. Історичні витрати в USD зафіксовані при внесенні."
        .Font.Italic = True
        .WrapText = True
        .RowHeight = 38
    End With
    Exit Sub
FooterFailed:
    originalNo = Err.Number
    originalText = Err.Description
    Debug.Print "TC_FINANCIAL_FOOTER_ERROR: " & footerStep & "; " & CStr(originalNo) & "; " & originalText
    Err.Raise vbObjectError + 3371, "FinancialFooter", _
        "Крок [" & footerStep & "]; первинна помилка " & CStr(originalNo) & ": " & originalText
End Sub


' v10.14: all monetary amounts come from the recorded USD ledger.
' Planned active-trip revenue is never presented as earned income.
Private Sub ActiveFinancialSnapshot(ByVal sh As Worksheet, ByVal atRow As Long, _
    ByVal trips As Worksheet, ByVal tripCosts As Object)
    Dim r As Long, tripId As String, nActive As Long
    Dim forecast As Double, spent As Double
    For r = ROW0 To Last(trips, 1)
        If S(trips.Cells(r, 8).Value2) = "У рейсі" Then
            tripId = S(trips.Cells(r, 1).Value2)
            If tripId <> "" Then
                nActive = nActive + 1
                forecast = forecast + N(trips.Cells(r, 14).Value2)
                spent = spent + Money(tripCosts, tripId)
            End If
        End If
    Next r
    Bar sh, atRow, "АКТИВНІ РЕЙСИ  |  планова виручка та вже зареєстровані витрати"
    TableLabels sh, atRow + 1, Array("Активних", "Дохід план USD", "Витрати факт USD", "Різниця USD", _
        "Як читати показник", "", "", "", "", "", "", "")
    sh.Cells(atRow + 2, 1).Value = nActive
    sh.Cells(atRow + 2, 2).Value = forecast
    sh.Cells(atRow + 2, 3).Value = spent
    sh.Cells(atRow + 2, 4).Value = forecast - spent
    sh.Cells(atRow + 2, 5).Value = "Прогноз, НЕ прибуток: майбутні витрати ще не відомі"
    Stripe sh, atRow + 2
    sh.Cells(atRow + 2, 5).WrapText = True
    sh.Rows(atRow + 2).RowHeight = 43
End Sub

Private Sub ActiveRiskSignals(ByVal sh As Worksheet, ByVal atRow As Long, _
    ByVal trips As Worksheet, ByVal tripCosts As Object, ByVal people As Object)
    Dim r As Long, target As Long, tripId As String, signal As String, tripAge As Long
    Dim cost As Double, forecast As Double, driverId As String
    Bar sh, atRow, "КОНТРОЛЬ АКТИВНИХ РЕЙСІВ  |  потенційні проблеми обліку"
    TableLabels sh, atRow + 1, Array("Рейс ID", "Водій", "Потрібна увага", _
        "План USD", "Витрати USD", "Днів у рейсі", "", "", "", "", "", "")
    target = atRow + 2
    For r = ROW0 To Last(trips, 1)
        If S(trips.Cells(r, 8).Value2) = "У рейсі" Then
            tripId = S(trips.Cells(r, 1).Value2)
            If tripId <> "" Then
                signal = ""
                tripAge = TripDays(trips.Cells(r, 9).Value, Empty, True)
                cost = Money(tripCosts, tripId)
                forecast = N(trips.Cells(r, 14).Value2)
                If Not IsNumeric(trips.Cells(r, 27).Value2) Or N(trips.Cells(r, 27).Value2) <= 0 Then _
                    signal = "Немає поточного одометра"
                If tripAge >= 14 Then
                    If signal <> "" Then signal = signal & "; "
                    signal = signal & "У рейсі 14+ днів"
                End If
                If forecast > 0 And cost > forecast Then
                    If signal <> "" Then signal = signal & "; "
                    signal = signal & "Витрати вищі за планову виручку"
                End If
                If signal <> "" Then
                    driverId = S(trips.Cells(r, 3).Value2)
                    sh.Cells(target, 1).Value = tripId
                    sh.Cells(target, 2).Value = Display(people, driverId)
                    sh.Cells(target, 3).Value = signal
                    sh.Cells(target, 4).Value = forecast
                    sh.Cells(target, 5).Value = cost
                    sh.Cells(target, 6).Value = tripAge
                    Stripe sh, target
                    sh.Cells(target, 3).WrapText = True
                    sh.Cells(target, 3).Font.Color = RGB(142, 66, 16)
                    sh.Rows(target).RowHeight = 45
                    target = target + 1
                End If
            End If
        End If
    Next r
    If target = atRow + 2 Then
        sh.Cells(target, 1).Value = "Сигналів немає"
        sh.Cells(target, 3).Value = "Перевірено поточний одометр, 14-денний поріг і перевищення планової виручки"
        Stripe sh, target
        sh.Cells(target, 3).WrapText = True
        sh.Rows(target).RowHeight = 42
    End If
End Sub

' Archive finance is grouped by the COMPLETION month of each trip.
' Its expense column includes ALL ledger expenses belonging to those closed trips,
' regardless of each receipt date. This is distinct from the PDF cash-date view.
Private Function ArchiveMonthlyTable(ByVal sh As Worksheet, ByVal atRow As Long, _
    ByVal trips As Worksheet, ByVal tripCosts As Object) As Long
    Dim counts As Object, inc As Object, costs As Object, paid As Object, km As Object
    Dim r As Long, i As Long, j As Long, writeRow As Long, tripId As String, monthId As String
    Dim keys As Variant, temp As Variant, totalRevenue As Double, totalExpense As Double
    Dim totalPaid As Double, totalKm As Double, totalCount As Double
    Set counts = CreateObject("Scripting.Dictionary")
    Set inc = CreateObject("Scripting.Dictionary")
    Set costs = CreateObject("Scripting.Dictionary")
    Set paid = CreateObject("Scripting.Dictionary")
    Set km = CreateObject("Scripting.Dictionary")
    Bar sh, atRow, "АРХІВ ПО МІСЯЦЯХ  |  віднесення рейсу і всіх його витрат до місяця завершення"
    TableLabels sh, atRow + 1, Array("Місяць", "Рейсів", "Дохід USD", "Витрати USD", _
        "Результат USD", "Оплачено USD", "Не оплачено USD", "Пробіг км", "Маржа %", _
        "", "", "")
    For r = ROW0 To Last(trips, 1)
        tripId = S(trips.Cells(r, 1).Value2)
        If tripId <> "" And S(trips.Cells(r, 8).Value2) = "Завершено" Then
            If IsDate(trips.Cells(r, 10).Value) Then
                monthId = Format$(CDate(trips.Cells(r, 10).Value), "yyyy-mm")
                AddMoney counts, monthId, 1
                AddMoney inc, monthId, N(trips.Cells(r, 14).Value2)
                AddMoney costs, monthId, Money(tripCosts, tripId)
                AddMoney paid, monthId, N(trips.Cells(r, 15).Value2)
                AddMoney km, monthId, N(trips.Cells(r, 13).Value2)
            End If
        End If
    Next r
    writeRow = atRow + 2
    If counts.Count = 0 Then
        sh.Cells(writeRow, 1).Value = "У журналі немає завершених рейсів із датою прибуття"
        Stripe sh, writeRow
        ArchiveMonthlyTable = 0
        Exit Function
    End If
    keys = counts.Keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CStr(keys(j)) < CStr(keys(i)) Then
                temp = keys(i): keys(i) = keys(j): keys(j) = temp
            End If
        Next j
    Next i
    For i = LBound(keys) To UBound(keys)
        monthId = CStr(keys(i))
        sh.Cells(writeRow, 1).Value = monthId
        sh.Cells(writeRow, 2).Value = Money(counts, monthId)
        sh.Cells(writeRow, 3).Value = Money(inc, monthId)
        sh.Cells(writeRow, 4).Value = Money(costs, monthId)
        sh.Cells(writeRow, 5).Value = Money(inc, monthId) - Money(costs, monthId)
        sh.Cells(writeRow, 6).Value = Money(paid, monthId)
        sh.Cells(writeRow, 7).Value = Money(inc, monthId) - Money(paid, monthId)
        sh.Cells(writeRow, 8).Value = Money(km, monthId)
        If Money(inc, monthId) > 0 Then _
            sh.Cells(writeRow, 9).Value = (Money(inc, monthId) - Money(costs, monthId)) * 100 / Money(inc, monthId)
        Stripe sh, writeRow
        totalRevenue = totalRevenue + Money(inc, monthId)
        totalExpense = totalExpense + Money(costs, monthId)
        totalPaid = totalPaid + Money(paid, monthId)
        totalKm = totalKm + Money(km, monthId)
        totalCount = totalCount + Money(counts, monthId)
        writeRow = writeRow + 1
    Next i
    ArchiveMonthlyTable = writeRow - 1
    With sh.Range(sh.Cells(writeRow, 1), sh.Cells(writeRow, 12))
        .Interior.Color = RGB(212, 234, 221)
        .Font.Bold = True
        .RowHeight = 31
    End With
    sh.Cells(writeRow, 1).Value = "РАЗОМ"
    sh.Cells(writeRow, 2).Value = totalCount
    sh.Cells(writeRow, 3).Value = totalRevenue
    sh.Cells(writeRow, 4).Value = totalExpense
    sh.Cells(writeRow, 5).Value = totalRevenue - totalExpense
    sh.Cells(writeRow, 6).Value = totalPaid
    sh.Cells(writeRow, 7).Value = totalRevenue - totalPaid
    sh.Cells(writeRow, 8).Value = totalKm
    If totalRevenue > 0 Then sh.Cells(writeRow, 9).Value = (totalRevenue - totalExpense) * 100 / totalRevenue
    With sh.Range(sh.Cells(writeRow + 1, 1), sh.Cells(writeRow + 1, 12))
        .Merge
        .Value = "Не оплачено = нараховано мінус зафіксована оплата. Витрати віднесено до місяця завершення рейсу, а не до дати квитанції."
        .Font.Italic = True
        .WrapText = True
        .RowHeight = 36
    End With
    On Error Resume Next
    sh.Range(sh.Cells(atRow + 2, 3), sh.Cells(writeRow, 8)).NumberFormat = "#,##0.00"
    sh.Range(sh.Cells(atRow + 2, 9), sh.Cells(writeRow, 9)).NumberFormat = "0.0"
    On Error GoTo 0
End Function

Private Sub ArchiveTrendChart(ByVal sh As Worksheet, ByVal monthlyHeader As Long, ByVal lastMonthRow As Long)
    Dim obj As ChartObject, topRow As Long
    If lastMonthRow < monthlyHeader + 2 Then Exit Sub
    On Error GoTo ChartNotSupported
    topRow = Application.Max(lastMonthRow + 5, Last(sh, 1) + 6)
    Set obj = sh.ChartObjects.Add(sh.Range("A" & topRow).Left, _
        sh.Range("A" & topRow).Top, 790, 275)
    With obj.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Завершені рейси: дохід і витрати за місяцями (USD)"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Дохід USD"
        .SeriesCollection(1).XValues = sh.Range("A" & (monthlyHeader + 2) & ":A" & lastMonthRow)
        .SeriesCollection(1).Values = sh.Range("C" & (monthlyHeader + 2) & ":C" & lastMonthRow)
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "Витрати USD"
        .SeriesCollection(2).XValues = sh.Range("A" & (monthlyHeader + 2) & ":A" & lastMonthRow)
        .SeriesCollection(2).Values = sh.Range("D" & (monthlyHeader + 2) & ":D" & lastMonthRow)
        .HasLegend = True
    End With
    Exit Sub
ChartNotSupported:
    ' Charts are optional. The verified ledger and monthly table remain available.
    Debug.Print "ArchiveTrendChart: " & Err.Number & " " & Err.Description
End Sub

Public Sub TC_RefreshBoth()
    Dim trips As Worksheet, expenses As Worksheet, live As Worksheet, history As Worksheet
    Dim persons As Object, paths As Object, byTrip As Object, categoryTrip As Object, stateOf As Object
    Dim byCategoryAll As Object, byCategoryFinished As Object
    Dim tripDriver As Object, driverCompleted As Object, driverActive As Object, driverKm As Object
    Dim driverDays As Object, driverIncome As Object, driverCostAll As Object, driverCostDone As Object
    Dim driverOverviewRow As Long, tripSectionRow As Long, key As Variant
    Dim tr As Long, er As Long, lastTrip As Long, id As String, category As String, state As String, driver As String
    Dim active As Long, completed As Long, planned As Long, days As Long, durationSum As Long, km As Double
    Dim activeLine As Long, archiveLine As Long, archiveStart As Long, detailStart As Long, detailLine As Long
    Dim categoriesEnd As Long, liveTotalsAt As Long
    Dim monthlyHeader As Long, lastMonthlyRow As Long, archiveFinanceEnd As Long
    Dim incomeFinished As Double, expenseAll As Double, expenseFinished As Double
    Dim planStart As Variant, actualEnd As Variant, cost As Double, total As Double
    Dim refreshStage As String, sourceError As Long, sourceText As String, errorSource As String
    Dim sourceLine As Long, errorText As String
    Dim oldUpdating As Boolean, oldEvents As Boolean, viewPrepared As Boolean
    Dim categoryText As Object, v As Variant, lookupKey As String, currentText As String
    Dim detailBuffer() As Variant, detailValues() As Variant
    Dim costRowCount As Long, maxExpenseRows As Long, colIndex As Long
    Dim detailRegion As Range
    refreshStage = "01 / sheets"
    On Error GoTo Failed
    oldUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    viewPrepared = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set trips = W("TC_Рейси"): Set expenses = W("TC_Витрати")
    Set live = W("Поточний стан"): Set history = W("Архів рейсів")
    refreshStage = "02 / driver and route catalogues"
    Set persons = Drivers(): Set paths = Routes()
    refreshStage = "03 / dictionary initialization"
    Set byTrip = CreateObject("Scripting.Dictionary")
    Set categoryTrip = CreateObject("Scripting.Dictionary")
    Set categoryText = CreateObject("Scripting.Dictionary")
    categoryText.CompareMode = vbTextCompare
    Set stateOf = CreateObject("Scripting.Dictionary")
    Set byCategoryAll = CreateObject("Scripting.Dictionary")
    Set byCategoryFinished = CreateObject("Scripting.Dictionary")
    byCategoryAll.CompareMode = vbTextCompare: byCategoryFinished.CompareMode = vbTextCompare
    stateOf.CompareMode = vbTextCompare
    byTrip.CompareMode = vbTextCompare: categoryTrip.CompareMode = vbTextCompare
    Set tripDriver = CreateObject("Scripting.Dictionary")
    Set driverCompleted = CreateObject("Scripting.Dictionary")
    Set driverActive = CreateObject("Scripting.Dictionary")
    Set driverKm = CreateObject("Scripting.Dictionary")
    Set driverDays = CreateObject("Scripting.Dictionary")
    Set driverIncome = CreateObject("Scripting.Dictionary")
    Set driverCostAll = CreateObject("Scripting.Dictionary")
    Set driverCostDone = CreateObject("Scripting.Dictionary")
    refreshStage = "04 / map trip ownership"
    tripDriver.CompareMode = vbTextCompare
    For tr = ROW0 To Last(trips, 1)
        id = S(trips.Cells(tr, 1).Value2)
        If id <> "" Then
            driver = S(trips.Cells(tr, 3).Value2)
            tripDriver(id) = driver
            stateOf(id) = S(trips.Cells(tr, 8).Value2)
            If driver <> "" And Not persons.Exists(driver) Then persons(driver) = driver
        End If
    Next tr
    refreshStage = "05 / aggregate recorded expenses"
    maxExpenseRows = Application.Max(0, Last(expenses, 2) - ROW0 + 1)
    If maxExpenseRows > 0 Then ReDim detailBuffer(1 To maxExpenseRows, 1 To 12)
    For er = ROW0 To Last(expenses, 2)
        id = S(expenses.Cells(er, 2).Value2)
        If id <> "" Then
            cost = N(expenses.Cells(er, 8).Value2)
            AddMoney byTrip, id, cost
            category = S(expenses.Cells(er, 3).Value2)
            If category = "" Then category = "Інші"
            AddMoney categoryTrip, id & "|" & category, cost
            AddMoney byCategoryAll, category, cost
            If tripDriver.Exists(id) Then AddMoney driverCostAll, CStr(tripDriver(id)), cost
            expenseAll = expenseAll + cost
            If stateOf.Exists(id) Then
                If CStr(stateOf(id)) = "Завершено" Then
                    AddMoney byCategoryFinished, category, cost
                    costRowCount = costRowCount + 1
                    detailBuffer(costRowCount, 1) = id
                    detailBuffer(costRowCount, 2) = expenses.Cells(er, 9).Value2
                    detailBuffer(costRowCount, 3) = expenses.Cells(er, 3).Value2
                    detailBuffer(costRowCount, 4) = expenses.Cells(er, 4).Value2
                    detailBuffer(costRowCount, 5) = expenses.Cells(er, 5).Value2
                    detailBuffer(costRowCount, 6) = expenses.Cells(er, 6).Value2
                    detailBuffer(costRowCount, 7) = expenses.Cells(er, 7).Value2
                    detailBuffer(costRowCount, 8) = expenses.Cells(er, 8).Value2
                    detailBuffer(costRowCount, 12) = expenses.Cells(er, 10).Value2
                End If
            End If
        End If
    Next er

    ' Build category strings once. Previous report scanned every category per trip.
    For Each v In categoryTrip.Keys
        lookupKey = CStr(v)
        tr = InStr(1, lookupKey, "|", vbBinaryCompare)
        If tr > 1 Then
            id = Left$(lookupKey, tr - 1)
            currentText = Mid$(lookupKey, tr + 1) & " " & Format$(Money(categoryTrip, lookupKey), "#,##0.00") & " USD"
            If categoryText.Exists(id) Then
                categoryText(id) = CStr(categoryText(id)) & "; " & currentText
            Else
                categoryText.Add id, currentText
            End If
        End If
    Next v
    refreshStage = "06 / render live header"
    Header live, "ПОТОЧНИЙ СТАН", "активні рейси та їх фактичні витрати"
    refreshStage = "07 / render archive header"
    Header history, "АРХІВ РЕЙСІВ", "завершені рейси; середня тривалість та структура витрат"
    refreshStage = "08 / company heading"
    Bar live, 5, "ОГЛЯД КОМПАНІЇ  |  автоматичне оновлення при відкритті аркуша"
    Bar history, 5, "ОГЛЯД АРХІВУ  |  лише завершені рейси"
    Bar live, 10, "ВОДІЇ КОМПАНІЇ  |  рейси, пробіг, дохід, витрати та рентабельність"
    Bar history, 10, "ЕФЕКТИВНІСТЬ ВОДІЇВ  |  виконані рейси та їх фінансові результати"
    ' Leave enough space for any number of drivers; never overwrite the trip table.
    refreshStage = "09 / section layout"
    monthlyHeader = Application.Max(20, 14 + persons.Count)
    refreshStage = "09a / archive monthly summary above details"
    lastMonthlyRow = ArchiveMonthlyTable(history, monthlyHeader, trips, byTrip)
    tripSectionRow = Application.Max(31, monthlyHeader + 10, lastMonthlyRow + 6)
    ActiveFinancialSnapshot live, Application.Max(20, 14 + persons.Count), trips, byTrip
    Bar live, tripSectionRow, "АКТИВНІ РЕЙСИ  |  фактичні дані станом на сьогодні"
    Bar history, tripSectionRow, "ЗАВЕРШЕНІ РЕЙСИ  |  повний перелік"
    TableLabels live, tripSectionRow + 2, Array("Рейс ID", "Водій", "Маршрут", "Авто", "Виїзд", "Днів", "Км план", "Км факт зараз", "Витрати USD", "Дохід план USD", "Статус", "На що витрачено (USD)")
    TableLabels history, tripSectionRow + 2, Array("Рейс ID", "Водій", "Маршрут", "Авто", "Виїзд", "Прибуття", "Днів", "Км факт", "Витрати USD", "Дохід USD", "Результат USD", "Структура витрат (USD)")
    refreshStage = "10 / trip details"
    activeLine = tripSectionRow + 3: archiveLine = tripSectionRow + 3
    For tr = ROW0 To Last(trips, 1)
        id = S(trips.Cells(tr, 1).Value2)
        If id <> "" Then
            state = S(trips.Cells(tr, 8).Value2)
            stateOf(id) = state
            driver = S(trips.Cells(tr, 3).Value2)
            planStart = trips.Cells(tr, 9).Value
            actualEnd = trips.Cells(tr, 10).Value
            Select Case state
                Case "Заплановано": planned = planned + 1
                Case "У рейсі"
                    active = active + 1
                    AddMoney driverActive, driver, 1
                    days = TripDays(planStart, actualEnd, True)
                    live.Cells(activeLine, 1).Value = id
                    live.Cells(activeLine, 2).Value = Display(persons, driver)
                    live.Cells(activeLine, 3).Value = Display(paths, S(trips.Cells(tr, 5).Value2))
                    live.Cells(activeLine, 4).Value = S(trips.Cells(tr, 4).Value2)
                    live.Cells(activeLine, 5).Value = planStart
                    live.Cells(activeLine, 6).Value = days
                    live.Cells(activeLine, 7).Value = N(trips.Cells(tr, 16).Value2)
                    If IsNumeric(trips.Cells(tr, 27).Value2) Then
                        live.Cells(activeLine, 8).Value = Application.Max(0, N(trips.Cells(tr, 27).Value2) - N(trips.Cells(tr, 11).Value2))
                    Else
                        live.Cells(activeLine, 8).Value = "одометр не внесено"
                    End If
                    live.Cells(activeLine, 9).Value = Money(byTrip, id)
                    live.Cells(activeLine, 10).Value = N(trips.Cells(tr, 14).Value2)
                    live.Cells(activeLine, 11).Value = state
                    live.Cells(activeLine, 12).Value = DisplayCategory(categoryText, id)
                    Stripe live, activeLine
                    live.Cells(activeLine, 12).WrapText = True
                    If Len(S(live.Cells(activeLine, 12).Value2)) > 85 Then live.Rows(activeLine).RowHeight = 52
                    activeLine = activeLine + 1
                Case "Завершено"
                    completed = completed + 1
                    days = TripDays(planStart, actualEnd, False)
                    durationSum = durationSum + days
                    expenseFinished = expenseFinished + Money(byTrip, id)
                    incomeFinished = incomeFinished + N(trips.Cells(tr, 14).Value2)
                    AddMoney driverCompleted, driver, 1
                    AddMoney driverDays, driver, days
                    AddMoney driverKm, driver, N(trips.Cells(tr, 13).Value2)
                    AddMoney driverIncome, driver, N(trips.Cells(tr, 14).Value2)
                    AddMoney driverCostDone, driver, Money(byTrip, id)
                    history.Cells(archiveLine, 1).Value = id
                    history.Cells(archiveLine, 2).Value = Display(persons, driver)
                    history.Cells(archiveLine, 3).Value = Display(paths, S(trips.Cells(tr, 5).Value2))
                    history.Cells(archiveLine, 4).Value = S(trips.Cells(tr, 4).Value2)
                    history.Cells(archiveLine, 5).Value = planStart
                    history.Cells(archiveLine, 6).Value = actualEnd
                    history.Cells(archiveLine, 7).Value = days
                    history.Cells(archiveLine, 8).Value = N(trips.Cells(tr, 13).Value2)
                    history.Cells(archiveLine, 9).Value = Money(byTrip, id)
                    history.Cells(archiveLine, 10).Value = N(trips.Cells(tr, 14).Value2)
                    history.Cells(archiveLine, 11).Value = N(trips.Cells(tr, 14).Value2) - Money(byTrip, id)
                    history.Cells(archiveLine, 12).Value = DisplayCategory(categoryText, id)
                    Stripe history, archiveLine
                    history.Cells(archiveLine, 12).WrapText = True
                    If Len(S(history.Cells(archiveLine, 12).Value2)) > 85 Then history.Rows(archiveLine).RowHeight = 52
                    archiveLine = archiveLine + 1
            End Select
        End If
    Next tr
    If active = 0 Then live.Cells(tripSectionRow + 3, 1).Value = "Активних рейсів немає": activeLine = tripSectionRow + 4
    If completed = 0 Then history.Cells(tripSectionRow + 3, 1).Value = "Завершених рейсів немає": archiveLine = tripSectionRow + 4
    refreshStage = "11 / driver overview live"
    DriverOverview live, persons, 12, driverCompleted, driverActive, driverKm, driverDays, driverIncome, driverCostAll, False
    refreshStage = "12 / driver overview archive"
    DriverOverview history, persons, 12, driverCompleted, driverActive, driverKm, driverDays, driverIncome, driverCostDone, True

    refreshStage = "13 / KPI current"
    KpiCard live, 1, "АКТИВНІ РЕЙСИ", CStr(active)
    KpiCard live, 4, "ЗАВЕРШЕНО", CStr(completed)
    KpiCard live, 7, "ЗАПЛАНОВАНО", CStr(planned)
    KpiCard live, 10, "ВИТРАТИ, USD", Format$(expenseAll, "#,##0.00")
    refreshStage = "14 / KPI archive"
    KpiCard history, 1, "ЗАВЕРШЕНО РЕЙСІВ", CStr(completed)
    If completed > 0 Then
        KpiCard history, 4, "СЕРЕДНЯ ТРИВАЛІСТЬ", Format$(durationSum / completed, "0.0") & " дн."
    Else
        KpiCard history, 4, "СЕРЕДНЯ ТРИВАЛІСТЬ", "—"
    End If
    KpiCard history, 7, "ДОХІД, USD", Format$(incomeFinished, "#,##0.00")
    KpiCard history, 10, "ВИТРАТИ, USD", Format$(expenseFinished, "#,##0.00")

    refreshStage = "15 / detailed archive layout"
    archiveStart = Application.Max(archiveLine + 2, tripSectionRow + 7)
    Bar history, archiveStart, "ПІДСУМОК АРХІВУ  |  завершено: " & completed & "  |  середня тривалість: " & IIf(completed > 0, Format$(durationSum / Application.Max(1, completed), "0.0"), "—") & " днів"
    detailStart = archiveStart + 3
    Bar history, detailStart, "ДЕТАЛІ ВИТРАТ У ЗАВЕРШЕНИХ РЕЙСАХ"
    TableLabels history, detailStart + 1, Array("Рейс ID", "Дата витрати", "Категорія", "Паркінг", "Сума", "Валюта", "Курс зафіксовано", "USD у журналі", "", "", "", "Коментар")
    refreshStage = "16 / detailed expense archive"
    detailLine = detailStart + 2
    For er = ROW0 To Last(expenses, 2)
        id = S(expenses.Cells(er, 2).Value2)
        If id <> "" Then
            If stateOf.Exists(id) Then
                If CStr(stateOf(id)) = "Завершено" Then
                    category = S(expenses.Cells(er, 3).Value2)
                    If category = "" Then category = "Інші"
                    AddMoney byCategoryFinished, category, N(expenses.Cells(er, 8).Value2)
                    history.Cells(detailLine, 1).Value = id
                    history.Cells(detailLine, 2).Value = expenses.Cells(er, 9).Value
                    history.Cells(detailLine, 3).Value = expenses.Cells(er, 3).Value2
                    history.Cells(detailLine, 4).Value = expenses.Cells(er, 4).Value2
                    history.Cells(detailLine, 5).Value = expenses.Cells(er, 5).Value2
                    history.Cells(detailLine, 6).Value = expenses.Cells(er, 6).Value2
                    history.Cells(detailLine, 7).Value = expenses.Cells(er, 7).Value2
                    history.Cells(detailLine, 8).Value = expenses.Cells(er, 8).Value2
                    history.Cells(detailLine, 12).Value = expenses.Cells(er, 10).Value2
                    Stripe history, detailLine
                    detailLine = detailLine + 1
                End If
            End If
        End If
    Next er
    ' Financial summary ends the archive AFTER all transaction records AND category subtotals.
    refreshStage = "17 / archive categories"
    categoriesEnd = CategoryTotalsTable(history, Application.Max(detailLine + 2, detailStart + 4), _
        "СТРУКТУРА ВИТРАТ  |  лише завершені рейси", byCategoryFinished)
    refreshStage = "18 / archive currency footer"
    FinancialFooter history, categoriesEnd + 2, "ФІНАНСИ ЗА ЗАВЕРШЕНИМИ РЕЙСАМИ", incomeFinished, expenseFinished
    archiveFinanceEnd = categoriesEnd + 14

    refreshStage = "19 / live totals layout"
    liveTotalsAt = Application.Max(activeLine + 2, tripSectionRow + 7)
    Bar live, liveTotalsAt, "ПІДСУМОК КОМПАНІЇ  |  зараз активних: " & active & "  |  виконано: " & completed & "  |  заплановано: " & planned
    refreshStage = "20 / live categories"
    categoriesEnd = CategoryTotalsTable(live, liveTotalsAt + 2, "ВИТРАТИ ЗА КАТЕГОРІЯМИ  |  усі зареєстровані рейси", byCategoryAll)
    refreshStage = "21 / live currency footer"
    FinancialFooter live, categoriesEnd + 2, "ДОХОДИ ЗАВЕРШЕНИХ РЕЙСІВ / ВСІ ЗАРЕЄСТРОВАНІ ВИТРАТИ", incomeFinished, expenseAll
    ' Further operational intelligence lives below the complete ledger tables.
    refreshStage = "22 / active financial snapshot already rendered near top
    refreshStage = "23 / active risk indicators"
    ActiveRiskSignals live, categoriesEnd + 20, trips, byTrip, persons
    refreshStage = "24 / archive monthly table already rendered near top"
    refreshStage = "25 / archive monthly chart"
    If lastMonthlyRow > 0 Then ArchiveTrendChart history, monthlyHeader, lastMonthlyRow

    ' Cosmetic formatting can fail under some Excel LTSC printers/configurations.
    ' Never discard correct ledger data and dashboard merely over a number format.
    refreshStage = "26 / presentation formats"
    On Error Resume Next
    live.Range("E" & (tripSectionRow + 3) & ":E" & Application.Max(activeLine, tripSectionRow + 3)).NumberFormat = "dd.mm.yyyy"
    history.Range("E" & (tripSectionRow + 3) & ":F" & Application.Max(archiveLine, tripSectionRow + 3)).NumberFormat = "dd.mm.yyyy"
    If detailLine > detailStart + 2 Then history.Range("B" & (detailStart + 2) & ":B" & (detailLine - 1)).NumberFormat = "dd.mm.yyyy"
    live.Columns("L").WrapText = True
    history.Columns("L").WrapText = True
    live.PageSetup.Orientation = xlLandscape
    history.PageSetup.Orientation = xlLandscape
    live.PageSetup.Zoom = False: live.PageSetup.FitToPagesWide = 1
    history.PageSetup.Zoom = False: history.PageSetup.FitToPagesWide = 1
    On Error GoTo Failed
    Exit Sub
Failed:
    ' Preserve original error stage and number; the previous handler hid the cause.
    sourceError = Err.Number
    sourceText = Err.Description
    errorSource = Err.Source
    sourceLine = Erl
    errorText = "ЕТАП " & refreshStage & "; VBA error " & CStr(sourceError) & _
        "; line " & CStr(sourceLine) & "; source " & errorSource & "; " & sourceText
    Debug.Print "TC_STATUS_FAILED: " & errorText
    On Error Resume Next
    ThisWorkbook.Worksheets("Поточний стан").Range("N1").Value = errorText
    ThisWorkbook.Worksheets("Архів рейсів").Range("N1").Value = errorText
    On Error GoTo 0
    Err.Raise vbObjectError + 3301, "TransportStatusFX.TC_RefreshBoth", errorText
End Sub

Public Sub TC_ShowLive()
    On Error GoTo Failed
    TC_RefreshBoth
    W("Поточний стан").Activate
    ActiveWindow.DisplayGridlines = False
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Поточний стан"
End Sub

Public Sub TC_ShowArchive()
    On Error GoTo Failed
    TC_RefreshBoth
    W("Архів рейсів").Activate
    ActiveWindow.DisplayGridlines = False
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Архів рейсів"
End Sub

Public Sub TC_ShowRates()
    W("Курси валют").Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Public Sub TC_RatesManualApply()
    Dim f As Worksheet, curr As Variant, r As Long, changed As Boolean, msg As String
    On Error GoTo Failed
    Set f = W("Курси валют")
    For r = 7 To 9
        curr = f.Cells(r, 2).Value2
        If Len(S(curr)) > 0 Then
            If Not IsNumeric(curr) Then Err.Raise vbObjectError + 3310, , "Курс у рядку " & r & " повинен бути числом."
            If CDbl(curr) <= 0 Then Err.Raise vbObjectError + 3311, , "Курс має бути більше нуля. Рядок " & r
            changed = True
        End If
    Next r
    If Not changed Then Err.Raise vbObjectError + 3312, , "Заповніть хоча б один із курсів USD, EUR, UAH у жовтих полях."
    If MsgBox("Застосувати введені курси PLN за одиницю валюти?", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    For r = 7 To 9
        If Len(S(f.Cells(r, 2).Value2)) > 0 Then
            f.Cells(r, 3).Value = Date
            f.Cells(r, 4).Value = "Вручну"
            f.Cells(r, 5).Value = Now
        End If
    Next r
    If Rate("USD") > 0 Then W("Налаштування").Range("B8").Value = Rate("USD")
    ThisWorkbook.Save
    TC_RefreshBoth
    MsgBox "Курси збережено. Інтернет-запиту не було.", vbInformation
    Exit Sub
Failed:
    MsgBox "Не вдалося зберегти курси: " & Err.Description, vbExclamation
End Sub

Private Function NbpRate(ByVal iso As String, ByRef effectiveDate As Date) As Double
    Dim req As Object, xml As Object, rateNode As Object, dateNode As Object, url As String
    Set req = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    req.setTimeouts 5000, 5000, 10000, 10000
    url = "https://api.nbp.pl/api/exchangerates/rates/a/" & LCase$(iso) & "/?format=xml"
    req.Open "GET", url, False
    req.Send
    If req.Status <> 200 Then Err.Raise vbObjectError + 3314, , "NBP: " & iso & " — HTTP " & req.Status
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False
    If Not xml.LoadXML(req.responseText) Then Err.Raise vbObjectError + 3315, , "NBP: некоректні XML-дані " & iso
    Set rateNode = xml.SelectSingleNode("//*[local-name()='Mid']")
    Set dateNode = xml.SelectSingleNode("//*[local-name()='EffectiveDate']")
    If rateNode Is Nothing Or dateNode Is Nothing Then Err.Raise vbObjectError + 3316, , "NBP: немає курсу або дати " & iso
    NbpRate = Val(Replace(CStr(rateNode.Text), ",", "."))
    If NbpRate <= 0 Then Err.Raise vbObjectError + 3317, , "NBP: недійсний курс " & iso
    effectiveDate = DateSerial(CInt(Left$(dateNode.Text, 4)), CInt(Mid$(dateNode.Text, 6, 2)), CInt(Mid$(dateNode.Text, 9, 2)))
End Function

Public Sub TC_RatesFetchNBP()
    Dim f As Worksheet, codes As Variant, rates(0 To 2) As Double, dates(0 To 2) As Date, i As Long
    On Error GoTo Failed
    If MsgBox("Завантажити курси USD, EUR та UAH до PLN з NBP? Тільки за цим натисканням буде Інтернет-запит.", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    codes = Array("USD", "EUR", "UAH")
    For i = 0 To 2
        rates(i) = NbpRate(CStr(codes(i)), dates(i))
    Next i
    Set f = W("Курси валют")
    For i = 0 To 2
        f.Cells(7 + i, 2).Value = rates(i)
        f.Cells(7 + i, 3).Value = dates(i)
        f.Cells(7 + i, 4).Value = "NBP (вручну кнопкою)"
        f.Cells(7 + i, 5).Value = Now
    Next i
    W("Налаштування").Range("B8").Value = rates(0)
    ThisWorkbook.Save
    TC_RefreshBoth
    MsgBox "Курси NBP USD/EUR/UAH застосовані. Автоматичного оновлення немає.", vbInformation
    Exit Sub
Failed:
    MsgBox "Курси не оновлено: " & Err.Description, vbExclamation
End Sub

Public Function TC_USDFrom(ByVal amount As Double, ByVal currency As String) As Double
    Dim baseUsd As Double, quote As Double
    currency = UCase$(Trim$(currency))
    If currency = "USD" Then TC_USDFrom = amount: Exit Function
    baseUsd = Rate("USD")
    quote = Rate(currency)
    If baseUsd <= 0 Or quote <= 0 Then Err.Raise vbObjectError + 3320, , "Для " & currency & " потрібні курси USD/PLN і " & currency & "/PLN на аркуші Курси валют."
    TC_USDFrom = Round(amount * quote / baseUsd, 2)
End Function

Public Function TC_PLNPerUnit(ByVal currency As String) As Double
    TC_PLNPerUnit = Rate(UCase$(Trim$(currency)))
End Function