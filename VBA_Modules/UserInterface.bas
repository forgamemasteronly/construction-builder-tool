' User Interface Module - Dashboard and Forms
' Handles the main dashboard and user input forms

Option Explicit

' ========== CONSTANTS ==========
Const DASHBOARD_SHEET = "Dashboard"
Const MATERIALS_SHEET = "Materials"
Const WORKERS_SHEET = "Workers"
Const PROJECT_SHEET = "Project"
Const SUMMARY_SHEET = "Summary"

' ========== DASHBOARD INITIALIZATION ==========

Sub InitializeDashboard()
    Dim ws As Worksheet
    Dim dashWs As Worksheet
    
    ' Create or get dashboard sheet
    On Error Resume Next
    Set dashWs = ThisWorkbook.Sheets(DASHBOARD_SHEET)
    If dashWs Is Nothing Then
        Set dashWs = ThisWorkbook.Sheets.Add
        dashWs.Name = DASHBOARD_SHEET
    End If
    On Error GoTo 0
    
    dashWs.Cells.Clear
    
    ' Setup dashboard layout
    Call SetupDashboardHeader(dashWs)
    Call SetupProjectSettings(dashWs)
    Call SetupMaterialSection(dashWs)
    Call SetupLaborSection(dashWs)
    Call SetupDimensionsSection(dashWs)
    Call SetupSummarySection(dashWs)
    
    MsgBox "Dashboard initialized successfully!", vbInformation
End Sub

Sub SetupDashboardHeader(ws As Worksheet)
    Dim titleRange As Range
    
    Set titleRange = ws.Range("A1:F1")
    titleRange.Merge
    titleRange.Value = "CONSTRUCTION PROJECT COST ESTIMATOR"
    titleRange.Font.Size = 16
    titleRange.Font.Bold = True
    titleRange.Font.Color = RGB(255, 255, 255)
    titleRange.Interior.Color = RGB(0, 51, 102)
    titleRange.HorizontalAlignment = xlCenter
    titleRange.VerticalAlignment = xlCenter
    titleRange.RowHeight = 30
    
    ws.Range("A2").Value = "Project Name:"
    ws.Range("B2").Name = "ProjectName"
    ws.Range("A3").Value = "Date Created:"
    ws.Range("B3").Value = Format(Now(), "yyyy-mm-dd")
    ws.Range("B3").Name = "ProjectDate"
End Sub

Sub SetupProjectSettings(ws As Worksheet)
    Dim row As Long
    row = 5
    
    ' Section header
    ws.Range("A" & row).Value = "PROJECT SETTINGS"
    ws.Range("A" & row & ":B" & row).Font.Bold = True
    ws.Range("A" & row & ":B" & row).Interior.Color = RGB(200, 200, 200)
    
    row = row + 1
    ws.Range("A" & row).Value = "Country:"
    ws.Range("B" & row).Value = "Philippines"
    ws.Range("B" & row).Name = "SelectedCountry"
    
    row = row + 1
    ws.Range("A" & row).Value = "Currency:"
    ws.Range("B" & row).Value = "PHP"
    ws.Range("B" & row).Name = "SelectedCurrency"
    
    row = row + 1
    ws.Range("A" & row).Value = "Exchange Rate (1 USD = ? PHP):"
    ws.Range("B" & row).Value = 56.5
    ws.Range("B" & row).Name = "ExchangeRate"
    
    row = row + 1
    ws.Range("A" & row).Value = "Last Updated:"
    ws.Range("B" & row).Value = Format(Now(), "yyyy-mm-dd hh:mm:ss")
    ws.Range("B" & row).Name = "RateLastUpdated"
End Sub

Sub SetupDimensionsSection(ws As Worksheet)
    Dim row As Long
    row = 11
    
    ' Section header
    ws.Range("A" & row).Value = "DIMENSIONS"
    ws.Range("A" & row & ":B" & row).Font.Bold = True
    ws.Range("A" & row & ":B" & row).Interior.Color = RGB(200, 200, 200)
    
    row = row + 1
    ws.Range("A" & row).Value = "Length (meters):"
    ws.Range("B" & row).Value = 0
    ws.Range("B" & row).Name = "DimensionLength"
    
    row = row + 1
    ws.Range("A" & row).Value = "Width (meters):"
    ws.Range("B" & row).Value = 0
    ws.Range("B" & row).Name = "DimensionWidth"
    
    row = row + 1
    ws.Range("A" & row).Value = "Height (meters):"
    ws.Range("B" & row).Value = 0
    ws.Range("B" & row).Name = "DimensionHeight"
    
    row = row + 1
    ws.Range("A" & row).Value = "Thickness (meters):"
    ws.Range("B" & row).Value = 0
    ws.Range("B" & row).Name = "DimensionThickness"
    
    row = row + 1
    ws.Range("A" & row).Value = "Area Calculated (sqm):"
    ws.Range("B" & row).Formula = "=DimensionLength*DimensionWidth"
    ws.Range("B" & row).Name = "CalculatedArea"
    
    row = row + 1
    ws.Range("A" & row).Value = "Volume Calculated (cum):"
    ws.Range("B" & row).Formula = "=DimensionLength*DimensionWidth*DimensionHeight"
    ws.Range("B" & row).Name = "CalculatedVolume"
End Sub

Sub SetupMaterialSection(ws As Worksheet)
    Dim row As Long
    row = 19
    
    ' Section header
    ws.Range("A" & row).Value = "MATERIALS"
    ws.Range("A" & row & ":D" & row).Font.Bold = True
    ws.Range("A" & row & ":D" & row).Interior.Color = RGB(200, 200, 200)
    
    row = row + 1
    ws.Range("A" & row).Value = "Material Name"
    ws.Range("B" & row).Value = "Quantity"
    ws.Range("C" & row).Value = "Unit Price"
    ws.Range("D" & row).Value = "Total Cost"
    
    ' Sample material rows (empty, to be filled in)
    Dim i As Long
    For i = 1 To 10
        row = row + 1
        ws.Range("A" & row).Name = "Material_Name_" & i
        ws.Range("B" & row).Name = "Material_Qty_" & i
        ws.Range("C" & row).Name = "Material_Price_" & i
        ws.Range("D" & row).Formula = "=IF(B" & row & "=0,0,B" & row & "*C" & row & ")"
    Next i
    
    row = row + 2
    ws.Range("A" & row).Value = "Total Materials Cost:"
    ws.Range("A" & row).Font.Bold = True
    ws.Range("B" & row).Formula = "=SUM(D20:D29)"
    ws.Range("B" & row).Name = "TotalMaterialCost"
    ws.Range("B" & row).Font.Bold = True
End Sub

Sub SetupLaborSection(ws As Worksheet)
    Dim row As Long
    row = 33
    
    ' Section header
    ws.Range("A" & row).Value = "LABOR"
    ws.Range("A" & row & ":D" & row).Font.Bold = True
    ws.Range("A" & row & ":D" & row).Interior.Color = RGB(200, 200, 200)
    
    row = row + 1
    ws.Range("A" & row).Value = "Worker Type"
    ws.Range("B" & row).Value = "Number of Workers"
    ws.Range("C" & row).Value = "Days Required"
    ws.Range("D" & row).Value = "Total Labor Cost"
    
    ' Sample worker rows (empty, to be filled in)
    Dim i As Long
    For i = 1 To 5
        row = row + 1
        ws.Range("A" & row).Name = "Worker_Type_" & i
        ws.Range("B" & row).Name = "Worker_Count_" & i
        ws.Range("C" & row).Name = "Worker_Days_" & i
        ws.Range("D" & row).Formula = "=IF(OR(B" & row & "=0,C" & row & "=0),0,B" & row & "*C" & row & "*250)"
    Next i
    
    row = row + 2
    ws.Range("A" & row).Value = "Total Labor Cost:"
    ws.Range("A" & row).Font.Bold = True
    ws.Range("B" & row).Formula = "=SUM(D35:D39)"
    ws.Range("B" & row).Name = "TotalLaborCost"
    ws.Range("B" & row).Font.Bold = True
End Sub

Sub SetupSummarySection(ws As Worksheet)
    Dim row As Long
    row = 43
    
    ' Section header
    ws.Range("A" & row).Value = "PROJECT SUMMARY"
    ws.Range("A" & row & ":B" & row).Font.Bold = True
    ws.Range("A" & row & ":B" & row).Interior.Color = RGB(0, 100, 0)
    ws.Range("A" & row & ":B" & row).Font.Color = RGB(255, 255, 255)
    
    row = row + 1
    ws.Range("A" & row).Value = "Total Materials Cost:"
    ws.Range("B" & row).Formula = "=TotalMaterialCost"
    ws.Range("B" & row).Name = "SummaryMaterialCost"
    ws.Range("B" & row).NumberFormat = "Currency"
    
    row = row + 1
    ws.Range("A" & row).Value = "Total Labor Cost:"
    ws.Range("B" & row).Formula = "=TotalLaborCost"
    ws.Range("B" & row).Name = "SummaryLaborCost"
    ws.Range("B" & row).NumberFormat = "Currency"
    
    row = row + 1
    ws.Range("A" & row).Value = "Contingency (10%):"
    ws.Range("B" & row).Formula = "=(SummaryMaterialCost+SummaryLaborCost)*0.10"
    ws.Range("B" & row).Name = "ContingencyCost"
    ws.Range("B" & row).NumberFormat = "Currency"
    
    row = row + 1
    ws.Range("A" & row).Value = "TOTAL PROJECT COST:"
    ws.Range("B" & row).Formula = "=SummaryMaterialCost+SummaryLaborCost+ContingencyCost"
    ws.Range("B" & row).Name = "TotalProjectCost"
    ws.Range("A" & row).Font.Bold = True
    ws.Range("A" & row).Font.Size = 12
    ws.Range("B" & row).Font.Bold = True
    ws.Range("B" & row).Font.Size = 12
    ws.Range("B" & row).Interior.Color = RGB(255, 255, 0)
    ws.Range("B" & row).NumberFormat = "Currency"
End Sub

' ========== USER INPUT FORMS ==========

Sub ShowMaterialInputForm()
    Dim ws As Worksheet
    Dim materialName As String
    Dim quantity As Double
    Dim unitPrice As Double
    Dim row As Long
    
    Set ws = ThisWorkbook.Sheets(DASHBOARD_SHEET)
    
    materialName = InputBox("Enter Material Name:", "Add Material")
    If materialName = "" Then Exit Sub
    
    quantity = CDbl(InputBox("Enter Quantity:", "Add Material"))
    If Err.Number <> 0 Then Exit Sub
    
    unitPrice = CDbl(InputBox("Enter Unit Price:", "Add Material"))
    If Err.Number <> 0 Then Exit Sub
    
    ' Find first empty row in materials section
    row = 20
    Do While ws.Range("A" & row).Value <> ""
        row = row + 1
    Loop
    
    If row <= 29 Then
        ws.Range("A" & row).Value = materialName
        ws.Range("B" & row).Value = quantity
        ws.Range("C" & row).Value = unitPrice
        MsgBox "Material added successfully!", vbInformation
    Else
        MsgBox "Materials section is full. Please remove some items.", vbExclamation
    End If
End Sub

Sub ShowWorkerInputForm()
    Dim ws As Worksheet
    Dim workerType As String
    Dim workerCount As Long
    Dim daysRequired As Double
    Dim row As Long
    
    Set ws = ThisWorkbook.Sheets(DASHBOARD_SHEET)
    
    workerType = InputBox("Enter Worker Specialization (e.g., Albañil, Electricista):", "Add Worker")
    If workerType = "" Then Exit Sub
    
    workerCount = CLng(InputBox("Enter Number of Workers:", "Add Worker"))
    If Err.Number <> 0 Then Exit Sub
    
    daysRequired = CDbl(InputBox("Enter Days Required:", "Add Worker"))
    If Err.Number <> 0 Then Exit Sub
    
    ' Find first empty row in labor section
    row = 35
    Do While ws.Range("A" & row).Value <> ""
        row = row + 1
    Loop
    
    If row <= 39 Then
        ws.Range("A" & row).Value = workerType
        ws.Range("B" & row).Value = workerCount
        ws.Range("C" & row).Value = daysRequired
        MsgBox "Worker added successfully!", vbInformation
    Else
        MsgBox "Workers section is full. Please remove some items.", vbExclamation
    End If
End Sub

Sub ShowProjectSettingsForm()
    Dim ws As Worksheet
    Dim projectName As String
    Dim country As String
    
    Set ws = ThisWorkbook.Sheets(DASHBOARD_SHEET)
    
    projectName = InputBox("Enter Project Name:", "Project Settings", ws.Range("ProjectName").Value)
    If projectName <> "" Then
        ws.Range("ProjectName").Value = projectName
    End If
    
    country = InputBox("Enter Country (Philippines or USA):", "Project Settings", ws.Range("SelectedCountry").Value)
    If country = "Philippines" Or country = "USA" Then
        ws.Range("SelectedCountry").Value = country
        If country = "Philippines" Then
            ws.Range("SelectedCurrency").Value = "PHP"
        Else
            ws.Range("SelectedCurrency").Value = "USD"
        End If
    ElseIf country <> "" Then
        MsgBox "Invalid country. Please enter 'Philippines' or 'USA'.", vbExclamation
    End If
End Sub

' ========== CALCULATION AND REPORTING ==========

Sub RecalculateProjectCost()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DASHBOARD_SHEET)
    
    ' Force recalculation
    ThisWorkbook.Calculate
    
    MsgBox "Project cost recalculated successfully!" & vbCrLf & _
           "Total Cost: " & Format(ws.Range("TotalProjectCost").Value, "Currency"), vbInformation
End Sub

Sub GenerateProjectReport()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim dashWs As Worksheet
    Dim row As Long
    
    Set dashWs = ThisWorkbook.Sheets(DASHBOARD_SHEET)
    
    ' Create or clear report sheet
    On Error Resume Next
    Set reportWs = ThisWorkbook.Sheets("ProjectReport")
    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Sheets.Add
        reportWs.Name = "ProjectReport"
    Else
        reportWs.Cells.Clear
    End If
    On Error GoTo 0
    
    row = 1
    
    ' Report header
    reportWs.Range("A" & row & ":B" & row).Merge
    reportWs.Range("A" & row).Value = "PROJECT COST REPORT"
    reportWs.Range("A" & row).Font.Size = 14
    reportWs.Range("A" & row).Font.Bold = True
    reportWs.Range("A" & row).Interior.Color = RGB(0, 51, 102)
    reportWs.Range("A" & row).Font.Color = RGB(255, 255, 255)
    
    row = row + 2
    reportWs.Range("A" & row).Value = "Project Information"
    reportWs.Range("A" & row).Font.Bold = True
    
    row = row + 1
    reportWs.Range("A" & row).Value = "Project Name:"
    reportWs.Range("B" & row).Value = dashWs.Range("ProjectName").Value
    
    row = row + 1
    reportWs.Range("A" & row).Value = "Country:"
    reportWs.Range("B" & row).Value = dashWs.Range("SelectedCountry").Value
    
    row = row + 1
    reportWs.Range("A" & row).Value = "Date Created:"
    reportWs.Range("B" & row).Value = dashWs.Range("ProjectDate").Value
    
    row = row + 2
    reportWs.Range("A" & row).Value = "Cost Summary"
    reportWs.Range("A" & row).Font.Bold = True
    
    row = row + 1
    reportWs.Range("A" & row).Value = "Materials Cost:"
    reportWs.Range("B" & row).Value = dashWs.Range("SummaryMaterialCost").Value
    reportWs.Range("B" & row).NumberFormat = "Currency"
    
    row = row + 1
    reportWs.Range("A" & row).Value = "Labor Cost:"
    reportWs.Range("B" & row).Value = dashWs.Range("SummaryLaborCost").Value
    reportWs.Range("B" & row).NumberFormat = "Currency"
    
    row = row + 1
    reportWs.Range("A" & row).Value = "Contingency (10%):"
    reportWs.Range("B" & row).Value = dashWs.Range("ContingencyCost").Value
    reportWs.Range("B" & row).NumberFormat = "Currency"
    
    row = row + 1
    reportWs.Range("A" & row).Value = "TOTAL PROJECT COST:"
    reportWs.Range("B" & row).Value = dashWs.Range("TotalProjectCost").Value
    reportWs.Range("A" & row).Font.Bold = True
    reportWs.Range("B" & row).Font.Bold = True
    reportWs.Range("A" & row).Interior.Color = RGB(255, 255, 0)
    reportWs.Range("B" & row).Interior.Color = RGB(255, 255, 0)
    reportWs.Range("B" & row).NumberFormat = "Currency"
    
    reportWs.Columns("A:B").AutoFit
    
    MsgBox "Project report generated successfully!", vbInformation
End Sub

Sub ExportToCSV()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fso As Object
    Dim textFile As Object
    
    Set ws = ThisWorkbook.Sheets(DASHBOARD_SHEET)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    filePath = ThisWorkbook.Path & "\ProjectEstimate_" & Format(Now(), "yyyymmdd_hhmmss") & ".csv"
    Set textFile = fso.CreateTextFile(filePath, True)
    
    ' Write project info
    textFile.WriteLine "PROJECT COST ESTIMATE"
    textFile.WriteLine ""
    textFile.WriteLine "Project Name," & ws.Range("ProjectName").Value
    textFile.WriteLine "Country," & ws.Range("SelectedCountry").Value
    textFile.WriteLine "Date Created," & ws.Range("ProjectDate").Value
    textFile.WriteLine ""
    textFile.WriteLine "COST SUMMARY"
    textFile.WriteLine "Materials Cost," & ws.Range("SummaryMaterialCost").Value
    textFile.WriteLine "Labor Cost," & ws.Range("SummaryLaborCost").Value
    textFile.WriteLine "Contingency (10%)," & ws.Range("ContingencyCost").Value
    textFile.WriteLine "TOTAL PROJECT COST," & ws.Range("TotalProjectCost").Value
    
    textFile.Close
    
    MsgBox "Project exported to: " & filePath, vbInformation
End Sub

Sub ClearAllData()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to clear all project data?", vbYesNo + vbExclamation)
    
    If response = vbYes Then
        Call InitializeDashboard
        MsgBox "All project data cleared!", vbInformation
    End If
End Sub
