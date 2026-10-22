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
    sheet.UsedRange.UnMerge
    sheet.Cells.Clear
    sheet.Cells.Font.Name = "Times New Roman"
    sheet.Cells.Font.Size = 11
    sheet.Cells.Font.Color = RGB(20, 48, 40)
    sheet.Columns("A").ColumnWidth = 23
    sheet.Columns("B").ColumnWidth = 26
    sheet.Columns("C").ColumnWidth = 40
    sheet.Columns("D").ColumnWidth = 18
    sheet.Columns("E").ColumnWidth = 17
    sheet.Columns("F").ColumnWidth = 17
    sheet.Columns("G").ColumnWidth = 14
    sheet.Columns("H").ColumnWidth = 17
    sheet.Columns("I").ColumnWidth = 20
    sheet.Columns("J").ColumnWidth = 20
    sheet.Columns("K").ColumnWidth = 21
    sheet.Columns("L").ColumnWidth = 45
    With sheet.Range("A1:L2")
        .Merge
        .Value = "ТРАНСПОРТНИЙ ЦЕНТР  /  " & kind
        .Interior.Color = RGB(12, 36, 31)
        .Font.Color = RGB(244, 234, 204)
        .Font.Size = 20
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .RowHeight = 26
    End With
    With sheet.Range("A3:L3")
        .Merge
        .Value = "Станом на " & Format$(Now, "dd.mm.yyyy HH:nn") & "  |  " & note
        .Font.Size = 10
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
        .Font.Size = 13
        .RowHeight = 25
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
        .RowHeight = 37
    End With
End Sub

Private Sub Stripe(ByVal sheet As Worksheet, ByVal r As Long)
    With sheet.Range(sheet.Cells(r, 1), sheet.Cells(r, 12))
        If r Mod 2 = 0 Then .Interior.Color = RGB(234, 243, 238) Else .Interior.Color = RGB(250, 251, 250)
        .RowHeight = 24
    End With
End Sub

Private Function Rate(ByVal iso As String) As Double
    Dim w As Worksheet, row As Long
    If iso = "PLN" Then Rate = 1: Exit Function
    Set w = W("Курси валют")
    Select Case UCase$(iso)
        Case "USD": row = 7
        Case "EUR": row = 8
        Case "UAH": row = 9
        Case Else: Exit Function
    End Select
    If IsNumeric(w.Cells(row, 2).Value2) Then
        If CDbl(w.Cells(row, 2).Value2) > 0 Then Rate = CDbl(w.Cells(row, 2).Value2)
    End If
End Function

' USD is the accounting ledger currency; all four visible totals are reference equivalents.
Private Function Equivalent(ByVal usdAmount As Double, ByVal currency As String) As String
    Dim usdRate As Double, destinationRate As Double
    If currency = "USD" Then Equivalent = Format$(usdAmount, "#,##0.00"): Exit Function
    usdRate = Rate("USD"): destinationRate = Rate(currency)
    If usdRate <= 0 Or destinationRate <= 0 Then
        Equivalent = "немає курсу"
    Else
        Equivalent = Format$(usdAmount * usdRate / destinationRate, "#,##0.00")
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

Private Sub FinancialFooter(ByVal sheet As Worksheet, ByVal line As Long, ByVal caption As String, ByVal revenueUSD As Double, ByVal costUSD As Double)
    Dim currencies As Variant, i As Long, fx As String
    Bar sheet, line, caption
    line = line + 1
    TableLabels sheet, line, Array("Валюта", "ДОХІД", "ВИТРАТИ", "РЕЗУЛЬТАТ", "Перерахунок", "", "", "", "", "", "", "")
    currencies = Array("PLN", "USD", "EUR", "UAH")
    For i = 0 To 3
        fx = CStr(currencies(i))
        sheet.Cells(line + 1 + i, 1).Value = fx
        sheet.Cells(line + 1 + i, 2).Value = Equivalent(revenueUSD, fx)
        sheet.Cells(line + 1 + i, 3).Value = Equivalent(costUSD, fx)
        sheet.Cells(line + 1 + i, 4).Value = Equivalent(revenueUSD - costUSD, fx)
        If fx = "USD" Then
            sheet.Cells(line + 1 + i, 5).Value = "Облікова валюта"
        Else
            sheet.Cells(line + 1 + i, 5).Value = "за курсом аркуша Курси валют"
        End If
        Stripe sheet, line + i + 1
        sheet.Cells(line + 1 + i, 1).Font.Bold = True
    Next i
    With sheet.Range(sheet.Cells(line + 6, 1), sheet.Cells(line + 6, 12))
        .Merge
        .Value = "Дохід: нарахування завершених рейсів; витрати: фактичні записи. Курс показує еквівалент, а не додатковий прибуток. Історичні витрати в USD зафіксовані при внесенні."
        .Font.Italic = True
        .WrapText = True
        .RowHeight = 38
    End With
End Sub

Public Sub TC_RefreshBoth()
    Dim trips As Worksheet, expenses As Worksheet, live As Worksheet, history As Worksheet
    Dim persons As Object, paths As Object, byTrip As Object, categoryTrip As Object, stateOf As Object
    Dim tr As Long, er As Long, lastTrip As Long, id As String, category As String, state As String, driver As String
    Dim active As Long, completed As Long, planned As Long, days As Long, durationSum As Long, km As Double
    Dim activeLine As Long, archiveLine As Long, archiveStart As Long, detailStart As Long, detailLine As Long
    Dim incomeFinished As Double, expenseAll As Double, expenseFinished As Double
    Dim planStart As Variant, actualEnd As Variant, cost As Double, total As Double
    On Error GoTo Failed
    Set trips = W("TC_Рейси"): Set expenses = W("TC_Витрати")
    Set live = W("Поточний стан"): Set history = W("Архів рейсів")
    Set persons = Drivers(): Set paths = Routes()
    Set byTrip = CreateObject("Scripting.Dictionary")
    Set categoryTrip = CreateObject("Scripting.Dictionary")
    Set stateOf = CreateObject("Scripting.Dictionary")
    stateOf.CompareMode = vbTextCompare
    byTrip.CompareMode = vbTextCompare: categoryTrip.CompareMode = vbTextCompare
    For er = ROW0 To Last(expenses, 2)
        id = S(expenses.Cells(er, 2).Value2)
        If id <> "" Then
            cost = N(expenses.Cells(er, 8).Value2)
            AddMoney byTrip, id, cost
            category = S(expenses.Cells(er, 3).Value2)
            If category = "" Then category = "Інші"
            AddMoney categoryTrip, id & "|" & category, cost
            expenseAll = expenseAll + cost
        End If
    Next er

    Header live, "ПОТОЧНИЙ СТАН", "активні рейси та їх фактичні витрати"
    Header history, "АРХІВ РЕЙСІВ", "завершені рейси; середня тривалість та структура витрат"
    Bar live, 5, "АКТИВНІ РЕЙСИ  |  факт станом на сьогодні"
    TableLabels live, 7, Array("Рейс ID", "Водій", "Маршрут", "Авто", "Виїзд", "Днів", "Км план", "Км факт зараз", "Витрати USD", "Дохід план USD", "Статус", "На що витрачено (USD)")
    Bar history, 5, "ЗАВЕРШЕНІ РЕЙСИ  |  повний архів"
    TableLabels history, 7, Array("Рейс ID", "Водій", "Маршрут", "Авто", "Виїзд", "Прибуття", "Днів", "Км факт", "Витрати USD", "Дохід USD", "Результат USD", "Структура витрат (USD)")
    activeLine = 8: archiveLine = 8
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
                    live.Cells(activeLine, 12).Value = CategorySummary(categoryTrip, id)
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
                    history.Cells(archiveLine, 12).Value = CategorySummary(categoryTrip, id)
                    Stripe history, archiveLine
                    history.Cells(archiveLine, 12).WrapText = True
                    If Len(S(history.Cells(archiveLine, 12).Value2)) > 85 Then history.Rows(archiveLine).RowHeight = 52
                    archiveLine = archiveLine + 1
            End Select
        End If
    Next tr
    If active = 0 Then live.Cells(8, 1).Value = "Активних рейсів немає": activeLine = 9
    If completed = 0 Then history.Cells(8, 1).Value = "Завершених рейсів немає": archiveLine = 9

    archiveStart = Application.Max(archiveLine + 2, 13)
    Bar history, archiveStart, "ПІДСУМОК АРХІВУ  |  завершено: " & completed & "  |  середня тривалість: " & IIf(completed > 0, Format$(durationSum / Application.Max(1, completed), "0.0"), "—") & " днів"
    detailStart = archiveStart + 3
    Bar history, detailStart, "ДЕТАЛІ ВИТРАТ У ЗАВЕРШЕНИХ РЕЙСАХ"
    TableLabels history, detailStart + 1, Array("Рейс ID", "Дата витрати", "Категорія", "Паркінг", "Сума", "Валюта", "Курс зафіксовано", "USD у журналі", "", "", "", "Коментар")
    detailLine = detailStart + 2
    For er = ROW0 To Last(expenses, 2)
        id = S(expenses.Cells(er, 2).Value2)
        If id <> "" Then
            If stateOf.Exists(id) Then
                If CStr(stateOf(id)) = "Завершено" Then
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
    ' Financial summary ends the archive AFTER all transaction records.
    FinancialFooter history, Application.Max(detailLine + 2, detailStart + 4), "ФІНАНСИ ЗА ЗАВЕРШЕНИМИ РЕЙСАМИ", incomeFinished, expenseFinished

    Bar live, Application.Max(activeLine + 2, 13), "ПІДСУМОК КОМПАНІЇ  |  зараз активних: " & active & "  |  виконано: " & completed & "  |  заплановано: " & planned
    FinancialFooter live, Application.Max(activeLine + 4, 15), "ДОХОДИ ЗАВЕРШЕНИХ РЕЙСІВ / ВСІ ЗАРЕЄСТРОВАНІ ВИТРАТИ", incomeFinished, expenseAll
    ' Cosmetic formatting can fail under some Excel LTSC printers/configurations.
    ' Never discard correct ledger data and dashboard merely over a number format.
    On Error Resume Next
    live.Range("E8:E" & Application.Max(activeLine, 8)).NumberFormat = "dd.mm.yyyy"
    history.Range("E8:F" & Application.Max(archiveLine, 8)).NumberFormat = "dd.mm.yyyy"
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
    Err.Raise vbObjectError + 3301, "TransportStatusFX.TC_RefreshBoth", "Не вдалося перерахувати стан та архів: " & Err.Description
End Sub

Public Sub TC_ShowLive()
    On Error GoTo Failed
    TC_RefreshBoth
    W("Поточний стан").Activate
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Поточний стан"
End Sub

Public Sub TC_ShowArchive()
    On Error GoTo Failed
    TC_RefreshBoth
    W("Архів рейсів").Activate
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Архів рейсів"
End Sub

Public Sub TC_ShowRates()
    W("Курси валют").Activate
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