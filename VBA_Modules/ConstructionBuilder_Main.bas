' Construction Builder Tool - Main VBA Module
' This module contains the core functionality for material and labor calculations

Option Explicit

' ========== CONSTANTS ==========
Const PHP_CURRENCY = "PHP"
Const USD_CURRENCY = "USD"
Const DEFAULT_WASTE_FACTOR = 0.12 ' 12% waste factor for materials
Const PHP_MIN_WAGE = 570 ' Daily minimum wage Philippines (example)
Const USD_MIN_WAGE = 96 ' Federal minimum wage per day (example)

' ========== DATA STRUCTURES ==========
Type MaterialItem
    Name As String
    Category As String
    Unit As String
    PricePerUnit As Double
    Quantity As Double
    Currency As String
End Type

Type WorkerType
    SpecializationPH As String
    SpecializationEN As String
    DailyRate As Double
    Currency As String
    Country As String
End Type

Type ProjectDimensions
    Length As Double
    Width As Double
    Height As Double
    Thickness As Double
    Unit As String ' "m" for meters, "ft" for feet
End Type

' ========== MAIN CALCULATION FUNCTIONS ==========

Function CalculateTotalMaterialCost(materials() As MaterialItem, exchangeRate As Double, country As String) As Double
    Dim totalCost As Double
    Dim i As Integer
    
    totalCost = 0
    
    For i = LBound(materials) To UBound(materials)
        totalCost = totalCost + CalculateMaterialCost(materials(i), exchangeRate, country)
    Next i
    
    CalculateTotalMaterialCost = totalCost
End Function

Function CalculateMaterialCost(material As MaterialItem, exchangeRate As Double, targetCountry As String) As Double
    Dim cost As Double
    Dim quantityWithWaste As Double
    
    ' Add waste factor
    quantityWithWaste = material.Quantity * (1 + DEFAULT_WASTE_FACTOR)
    
    ' Calculate base cost
    cost = quantityWithWaste * material.PricePerUnit
    
    ' Convert currency if needed
    If material.Currency <> targetCountry Then
        If material.Currency = PHP_CURRENCY And targetCountry = USD_CURRENCY Then
            cost = cost / exchangeRate
        ElseIf material.Currency = USD_CURRENCY And targetCountry = PHP_CURRENCY Then
            cost = cost * exchangeRate
        End If
    End If
    
    CalculateMaterialCost = cost
End Function

Function CalculateTotalLaborCost(workers As Collection, daysRequired As Double, country As String) As Double
    Dim totalCost As Double
    Dim i As Integer
    Dim worker As WorkerType
    
    totalCost = 0
    
    For i = 1 To workers.Count
        Set worker = workers(i)
        totalCost = totalCost + CalculateWorkerCost(worker, daysRequired, country)
    Next i
    
    CalculateTotalLaborCost = totalCost
End Function

Function CalculateWorkerCost(worker As WorkerType, daysRequired As Double, country As String) As Double
    Dim dailyRate As Double
    Dim totalCost As Double
    Dim overtimeDays As Double
    Dim overtimeRate As Double
    
    dailyRate = GetAdjustedDailyRate(worker.DailyRate, country, worker.Currency)
    
    ' Calculate overtime based on country labor laws
    If country = "Philippines" Then
        ' Philippine labor law: 8 hours standard, overtime after
        totalCost = dailyRate * daysRequired
        ' Add 13th month pay consideration (1/12 of annual)
        totalCost = totalCost * 1.083
    ElseIf country = "USA" Then
        ' US labor law: 40 hours/week standard, 1.5x after 40 hours
        ' Assuming 8-hour workday
        If daysRequired > 5 Then
            overtimeDays = daysRequired - 5
            totalCost = (5 * dailyRate) + (overtimeDays * dailyRate * 1.5)
        Else
            totalCost = daysRequired * dailyRate
        End If
    End If
    
    CalculateWorkerCost = totalCost
End Function

Function GetAdjustedDailyRate(baseRate As Double, country As String, currency As String) As Double
    Dim minWage As Double
    Dim adjustedRate As Double
    
    ' Get minimum wage for country
    If country = "Philippines" Then
        minWage = PHP_MIN_WAGE
    ElseIf country = "USA" Then
        minWage = USD_MIN_WAGE
    End If
    
    ' Ensure rate meets minimum wage
    If baseRate < minWage Then
        adjustedRate = minWage
    Else
        adjustedRate = baseRate
    End If
    
    GetAdjustedDailyRate = adjustedRate
End Function

Function ConvertCurrency(amount As Double, fromCurrency As String, toCurrency As String, exchangeRate As Double) As Double
    Dim convertedAmount As Double
    
    If fromCurrency = toCurrency Then
        ConvertCurrency = amount
    ElseIf fromCurrency = PHP_CURRENCY And toCurrency = USD_CURRENCY Then
        ConvertCurrency = amount / exchangeRate
    ElseIf fromCurrency = USD_CURRENCY And toCurrency = PHP_CURRENCY Then
        ConvertCurrency = amount * exchangeRate
    Else
        ConvertCurrency = amount
    End If
End Function

Function CalculateProjectCost(materials As Collection, workers As Collection, daysRequired As Double, country As String, exchangeRate As Double) As Double
    Dim materialCost As Double
    Dim laborCost As Double
    Dim totalCost As Double
    Dim i As Integer
    Dim material As MaterialItem
    
    ' Calculate material cost
    materialCost = 0
    For i = 1 To materials.Count
        Set material = materials(i)
        materialCost = materialCost + CalculateMaterialCost(material, exchangeRate, country)
    Next i
    
    ' Calculate labor cost
    laborCost = CalculateTotalLaborCost(workers, daysRequired, country)
    
    ' Total project cost
    totalCost = materialCost + laborCost
    
    CalculateProjectCost = totalCost
End Function

' ========== WORKER CLASSIFICATION ==========

Function GetWorkerSpecializations() As String()
    Dim specializations(10) As String
    
    specializations(0) = "Armalador (Steel Binder/Reinforcement Worker)"
    specializations(1) = "Albañil (Mason/Brick Layer)"
    specializations(2) = "Tindero (Carpenter/Woodworker)"
    specializations(3) = "Electricista (Electrician)"
    specializations(4) = "Plomero (Plumber)"
    specializations(5) = "Magamit (Equipment Operator)"
    specializations(6) = "Manggagawa (General Laborer)"
    specializations(7) = "Pintero (Painter)"
    specializations(8) = "Weldes (Welder)"
    specializations(9) = "Tile Layer (Tile Setter/Mason)"
    
    GetWorkerSpecializations = specializations
End Function

Function GetWorkerRate(specialization As String, country As String) As Double
    Dim rate As Double
    
    Select Case specialization
        Case "Armalador"
            rate = IIf(country = "Philippines", 650, 120)
        Case "Albañil"
            rate = IIf(country = "Philippines", 700, 130)
        Case "Tindero"
            rate = IIf(country = "Philippines", 680, 125)
        Case "Electricista"
            rate = IIf(country = "Philippines", 800, 150)
        Case "Plomero"
            rate = IIf(country = "Philippines", 750, 140)
        Case "Magamit"
            rate = IIf(country = "Philippines", 900, 160)
        Case "Manggagawa"
            rate = IIf(country = "Philippines", 570, 96)
        Case "Pintero"
            rate = IIf(country = "Philippines", 650, 120)
        Case "Weldes"
            rate = IIf(country = "Philippines", 850, 155)
        Case "Tile Layer"
            rate = IIf(country = "Philippines", 700, 130)
        Case Else
            rate = IIf(country = "Philippines", 570, 96)
    End Select
    
    GetWorkerRate = rate
End Function

' ========== DIMENSION CALCULATIONS ==========

Function CalculateArea(length As Double, width As Double, unit As String) As Double
    CalculateArea = length * width
End Function

Function CalculateVolume(length As Double, width As Double, height As Double, unit As String) As Double
    CalculateVolume = length * width * height
End Function

Function ConvertUnit(value As Double, fromUnit As String, toUnit As String) As Double
    Dim convertedValue As Double
    
    ' Convert to meters first
    If fromUnit = "ft" Then
        value = value * 0.3048
    End If
    
    ' Convert from meters to target unit
    If toUnit = "ft" Then
        convertedValue = value / 0.3048
    Else
        convertedValue = value
    End If
    
    ConvertUnit = convertedValue
End Function

' ========== MATERIAL CALCULATION ==========

Function CalculateMaterialQuantity(area As Double, thickness As Double, MaterialType As String) As Double
    Dim quantity As Double
    
    Select Case MaterialType
        Case "Concrete"
            quantity = area * thickness
        Case "Bricks"
            ' Assuming standard brick size
            quantity = (area / 0.06) ' approximately 17 bricks per sq meter
        Case "Tiles"
            ' Assuming standard tile size
            quantity = area / 0.25 ' approximately 40 tiles per sq meter
        Case "Paint"
            ' Assuming coverage of 10 sq meters per liter
            quantity = area / 10
        Case "Lumber"
            quantity = area * thickness
        Case Else
            quantity = area
    End Select
    
    CalculateMaterialQuantity = quantity
End Function

' ========== EXCHANGE RATE FUNCTIONS ==========

Function FetchExchangeRate() As Double
    ' This would integrate with an API to fetch current PHP to USD rate
    ' For now, returning a static example value
    ' In production, use XMLHTTP or similar to call a real API
    FetchExchangeRate = 56.5 ' Example: 1 USD = 56.5 PHP
End Function

Function GetExchangeRateFromSheet(ws As Worksheet) As Double
    ' Fetch exchange rate from a designated cell in the worksheet
    On Error Resume Next
    GetExchangeRateFromSheet = ws.Range("ExchangeRate").Value
    On Error GoTo 0
End Function

' ========== UTILITY FUNCTIONS ==========

Function RoundToDecimal(value As Double, decimals As Integer) As Double
    RoundToDecimal = Round(value, decimals)
End Function

Function FormatCurrency(amount As Double, currency As String) As String
    If currency = PHP_CURRENCY Then
        FormatCurrency = "₱" & Format(amount, "#,##0.00")
    ElseIf currency = USD_CURRENCY Then
        FormatCurrency = "$" & Format(amount, "#,##0.00")
    End If
End Function

Sub UpdateAllCalculations()
    ' Main sub to update all calculations in the workbook
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ' Recalculate all formulas
    ThisWorkbook.Calculate
    
    MsgBox "All calculations updated successfully!", vbInformation
End Sub
