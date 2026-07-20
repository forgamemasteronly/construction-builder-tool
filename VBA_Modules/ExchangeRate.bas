' Exchange Rate Module - Currency Conversion and Management
' Handles PHP to USD conversion and vice versa

Option Explicit

' ========== CONSTANTS ==========
Const EXCHANGE_RATE_API = "https://api.exchangerate-api.com/v4/latest/"
Const CACHE_DURATION = 3600 ' Cache exchange rates for 1 hour (in seconds)

' ========== EXCHANGE RATE STRUCTURES ==========
Type ExchangeRateData
    FromCurrency As String
    ToCurrency As String
    Rate As Double
    LastUpdated As Date
    Source As String
End Type

' ========== MAIN EXCHANGE RATE FUNCTIONS ==========

Function GetCurrentExchangeRate(fromCurrency As String, toCurrency As String) As Double
    Dim rate As Double
    Dim xmlhttp As Object
    Dim response As String
    Dim cacheRate As Double
    
    On Error GoTo ErrorHandler
    
    ' Check if we have a cached rate
    cacheRate = GetCachedExchangeRate(fromCurrency, toCurrency)
    If cacheRate > 0 Then
        GetCurrentExchangeRate = cacheRate
        Exit Function
    End If
    
    ' Fetch from API
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP.6.0")
    
    xmlhttp.Open "GET", EXCHANGE_RATE_API & fromCurrency, False
    xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    xmlhttp.Send
    
    If xmlhttp.Status = 200 Then
        response = xmlhttp.ResponseText
        rate = ParseExchangeRateResponse(response, toCurrency)
        
        ' Cache the rate
        Call CacheExchangeRate(fromCurrency, toCurrency, rate)
        
        GetCurrentExchangeRate = rate
    Else
        ' Fallback to manual rate
        GetCurrentExchangeRate = GetManualExchangeRate(fromCurrency, toCurrency)
    End If
    
    Exit Function
ErrorHandler:
    ' If API fails, use manual rate
    GetCurrentExchangeRate = GetManualExchangeRate(fromCurrency, toCurrency)
End Function

Function ParseExchangeRateResponse(response As String, toCurrency As String) As Double
    ' Parse JSON response from exchange rate API
    Dim rate As Double
    Dim searchString As String
    Dim startPos As Long
    Dim endPos As Long
    
    searchString = """" & toCurrency & """:" ' Looking for "USD": format
    startPos = InStr(response, searchString)
    
    If startPos > 0 Then
        startPos = InStr(startPos, response, ":")
        endPos = InStr(startPos, response, ",")
        If endPos = 0 Then
            endPos = InStr(startPos, response, "}")
        End If
        
        rate = CDbl(Trim(Mid(response, startPos + 1, endPos - startPos - 1)))
        ParseExchangeRateResponse = rate
    Else
        ParseExchangeRateResponse = 0
    End If
End Function

Function GetManualExchangeRate(fromCurrency As String, toCurrency As String) As Double
    Dim rate As Double
    
    ' Standard exchange rate reference (as of last update)
    ' These should be updated regularly
    If fromCurrency = "PHP" And toCurrency = "USD" Then
        rate = 1 / 56.5 ' PHP to USD (example)
    ElseIf fromCurrency = "USD" And toCurrency = "PHP" Then
        rate = 56.5 ' USD to PHP (example)
    ElseIf fromCurrency = toCurrency Then
        rate = 1
    Else
        rate = 1 ' Fallback
    End If
    
    GetManualExchangeRate = rate
End Function

Function GetExchangeRateFromSheet(ws As Worksheet) As Double
    ' Retrieve exchange rate from designated worksheet cell
    Dim rate As Double
    
    On Error Resume Next
    rate = ws.Range("ExchangeRate").Value
    On Error GoTo 0
    
    If rate = 0 Then
        rate = GetCurrentExchangeRate("USD", "PHP")
    End If
    
    GetExchangeRateFromSheet = rate
End Function

' ========== CURRENCY CONVERSION FUNCTIONS ==========

Function ConvertPHPtoUSD(amountPHP As Double, exchangeRate As Double) As Double
    If exchangeRate <= 0 Then
        exchangeRate = GetCurrentExchangeRate("USD", "PHP")
    End If
    ConvertPHPtoUSD = amountPHP / exchangeRate
End Function

Function ConvertUSDtoPHP(amountUSD As Double, exchangeRate As Double) As Double
    If exchangeRate <= 0 Then
        exchangeRate = GetCurrentExchangeRate("USD", "PHP")
    End If
    ConvertUSDtoPHP = amountUSD * exchangeRate
End Function

Function ConvertCurrencyAmount(amount As Double, fromCurrency As String, toCurrency As String, exchangeRate As Double) As Double
    Dim convertedAmount As Double
    
    If fromCurrency = toCurrency Then
        ConvertCurrencyAmount = amount
        Exit Function
    End If
    
    If exchangeRate <= 0 Then
        exchangeRate = GetCurrentExchangeRate(fromCurrency, toCurrency)
    End If
    
    Select Case fromCurrency & "-" & toCurrency
        Case "PHP-USD"
            convertedAmount = amount / exchangeRate
        Case "USD-PHP"
            convertedAmount = amount * exchangeRate
        Case Else
            convertedAmount = amount
    End Select
    
    ConvertCurrencyAmount = convertedAmount
End Function

Function ApplyExchangeRateToPrices(ws As Worksheet, fromCurrency As String, toCurrency As String, exchangeRate As Double) As Boolean
    ' Apply exchange rate conversion to all prices in a worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim priceCell As Range
    Dim newPrice As Double
    
    On Error GoTo ErrorHandler
    
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row ' Assuming prices are in column D
    
    For i = 2 To lastRow ' Skip header row
        If ws.Range("D" & i).Value <> "" Then
            newPrice = ConvertCurrencyAmount(ws.Range("D" & i).Value, fromCurrency, toCurrency, exchangeRate)
            ws.Range("D" & i).Value = newPrice
        End If
    Next i
    
    ApplyExchangeRateToPrices = True
    Exit Function
ErrorHandler:
    MsgBox "Error applying exchange rate: " & Err.Description, vbCritical
    ApplyExchangeRateToPrices = False
End Function

' ========== CACHING FUNCTIONS ==========

Function GetCachedExchangeRate(fromCurrency As String, toCurrency As String) As Double
    Dim ws As Worksheet
    Dim cacheRange As Range
    Dim lastRow As Long
    Dim i As Long
    Dim cachedRate As Double
    Dim lastUpdated As Date
    Dim cacheAge As Long
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("ExchangeRates")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Range("A" & i).Value = fromCurrency And ws.Range("B" & i).Value = toCurrency Then
            lastUpdated = ws.Range("D" & i).Value
            cacheAge = DateDiff("s", lastUpdated, Now())
            
            If cacheAge < CACHE_DURATION Then
                cachedRate = ws.Range("C" & i).Value
                GetCachedExchangeRate = cachedRate
                Exit Function
            End If
        End If
    Next i
    
    GetCachedExchangeRate = 0
    Exit Function
ErrorHandler:
    GetCachedExchangeRate = 0
End Function

Sub CacheExchangeRate(fromCurrency As String, toCurrency As String, rate As Double)
    Dim ws As Worksheet
    Dim lastRow As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ExchangeRates")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "ExchangeRates"
        ws.Range("A1").Value = "From"
        ws.Range("B1").Value = "To"
        ws.Range("C1").Value = "Rate"
        ws.Range("D1").Value = "LastUpdated"
    End If
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ws.Range("A" & lastRow).Value = fromCurrency
    ws.Range("B" & lastRow).Value = toCurrency
    ws.Range("C" & lastRow).Value = rate
    ws.Range("D" & lastRow).Value = Now()
End Sub

' ========== USER INTERFACE FUNCTIONS ==========

Sub UpdateExchangeRate()
    Dim rate As Double
    Dim ws As Worksheet
    
    rate = GetCurrentExchangeRate("USD", "PHP")
    
    If rate > 0 Then
        Set ws = ThisWorkbook.Sheets("Dashboard")
        ws.Range("ExchangeRate").Value = rate
        MsgBox "Exchange Rate Updated: 1 USD = " & Format(rate, "0.00") & " PHP", vbInformation
    Else
        MsgBox "Failed to update exchange rate. Please check your internet connection.", vbExclamation
    End If
End Sub

Sub SetManualExchangeRate()
    Dim inputRate As Variant
    Dim ws As Worksheet
    
    inputRate = InputBox("Enter PHP to USD exchange rate (e.g., 56.5):", "Manual Exchange Rate")
    
    If inputRate <> "" Then
        If IsNumeric(inputRate) Then
            Set ws = ThisWorkbook.Sheets("Dashboard")
            ws.Range("ExchangeRate").Value = CDbl(inputRate)
            MsgBox "Exchange rate set to: " & inputRate, vbInformation
        Else
            MsgBox "Invalid input. Please enter a valid number.", vbExclamation
        End If
    End If
End Sub

Function IsNumeric(text As String) As Boolean
    On Error Resume Next
    IsNumeric = Not IsNull(CDbl(text))
End Function

' ========== REPORTING FUNCTIONS ==========

Sub GenerateExchangeRateReport()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim row As Long
    
    Set ws = ThisWorkbook.Sheets("ExchangeRates")
    
    ' Create or clear report sheet
    On Error Resume Next
    Set reportWs = ThisWorkbook.Sheets("ExchangeRateReport")
    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Sheets.Add
        reportWs.Name = "ExchangeRateReport"
    Else
        reportWs.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Add headers
    reportWs.Range("A1").Value = "From Currency"
    reportWs.Range("B1").Value = "To Currency"
    reportWs.Range("C1").Value = "Exchange Rate"
    reportWs.Range("D1").Value = "Last Updated"
    reportWs.Range("E1").Value = "Age (Hours)"
    
    ' Copy and process data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    row = 2
    
    For i = 2 To lastRow
        reportWs.Range("A" & row).Value = ws.Range("A" & i).Value
        reportWs.Range("B" & row).Value = ws.Range("B" & i).Value
        reportWs.Range("C" & row).Value = ws.Range("C" & i).Value
        reportWs.Range("D" & row).Value = ws.Range("D" & i).Value
        reportWs.Range("E" & row).Value = Format((Now() - ws.Range("D" & i).Value) * 24, "0.00")
        row = row + 1
    Next i
    
    ' Format report
    reportWs.Columns("A:E").AutoFit
    With reportWs.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    MsgBox "Exchange Rate Report generated successfully!", vbInformation
End Sub

' ========== HISTORICAL TRACKING ==========

Sub LogExchangeRateHistory(fromCurrency As String, toCurrency As String, rate As Double)
    Dim ws As Worksheet
    Dim lastRow As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("RateHistory")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "RateHistory"
        ws.Range("A1").Value = "Date"
        ws.Range("B1").Value = "From"
        ws.Range("C1").Value = "To"
        ws.Range("D1").Value = "Rate"
    End If
    On Error GoTo 0
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ws.Range("A" & lastRow).Value = Now()
    ws.Range("B" & lastRow).Value = fromCurrency
    ws.Range("C" & lastRow).Value = toCurrency
    ws.Range("D" & lastRow).Value = rate
End Sub

Function GetAverageExchangeRate(fromCurrency As String, toCurrency As String, daysBack As Integer) As Double
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sum As Double
    Dim count As Integer
    Dim cutoffDate As Date
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("RateHistory")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    cutoffDate = DateAdd("d", -daysBack, Now())
    
    sum = 0
    count = 0
    
    For i = 2 To lastRow
        If ws.Range("B" & i).Value = fromCurrency And ws.Range("C" & i).Value = toCurrency Then
            If ws.Range("A" & i).Value >= cutoffDate Then
                sum = sum + ws.Range("D" & i).Value
                count = count + 1
            End If
        End If
    Next i
    
    If count > 0 Then
        GetAverageExchangeRate = sum / count
    Else
        GetAverageExchangeRate = 0
    End If
    
    Exit Function
ErrorHandler:
    GetAverageExchangeRate = 0
End Function
