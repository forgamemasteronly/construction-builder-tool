' Utility Module - Helper Functions and System Operations
' Contains shared utility functions used across modules

Option Explicit

' ========== STRING AND TEXT UTILITIES ==========

Function TrimString(inputText As String) As String
    TrimString = Trim(inputText)
End Function

Function ToUpperCase(inputText As String) As String
    ToUpperCase = UCase(inputText)
End Function

Function ToLowerCase(inputText As String) As String
    ToLowerCase = LCase(inputText)
End Function

Function ReplaceText(inputText As String, oldText As String, newText As String) As String
    ReplaceText = Replace(inputText, oldText, newText)
End Function

' ========== NUMBER UTILITIES ==========

Function RoundNumber(value As Double, decimals As Integer) As Double
    RoundNumber = Round(value, decimals)
End Function

Function FormatNumberToCurrency(amount As Double, currency As String) As String
    If currency = "PHP" Then
        FormatNumberToCurrency = "₱" & Format(amount, "#,##0.00")
    ElseIf currency = "USD" Then
        FormatNumberToCurrency = "$" & Format(amount, "#,##0.00")
    Else
        FormatNumberToCurrency = Format(amount, "#,##0.00")
    End If
End Function

Function CalculatePercentage(value As Double, percentage As Double) As Double
    CalculatePercentage = (value * percentage) / 100
End Function

Function AddPercentageMarkup(basePrice As Double, markupPercent As Double) As Double
    AddPercentageMarkup = basePrice * (1 + (markupPercent / 100))
End Function

Function CalculateDiscount(basePrice As Double, discountPercent As Double) As Double
    CalculateDiscount = basePrice * (1 - (discountPercent / 100))
End Function

' ========== ARRAY AND COLLECTION UTILITIES ==========

Function GetArrayLength(arr() As Variant) As Long
    On Error Resume Next
    GetArrayLength = UBound(arr) - LBound(arr) + 1
    On Error GoTo 0
End Function

Function IsInCollection(col As Collection, item As Variant) As Boolean
    Dim i As Long
    IsInCollection = False
    
    On Error Resume Next
    For i = 1 To col.Count
        If col(i) = item Then
            IsInCollection = True
            Exit Function
        End If
    Next i
End Function

Sub PrintCollection(col As Collection)
    Dim i As Long
    Dim output As String
    
    output = "Collection Contents:" & vbCrLf
    
    For i = 1 To col.Count
        output = output & i & ": " & col(i) & vbCrLf
    Next i
    
    MsgBox output
End Sub

' ========== DATE AND TIME UTILITIES ==========

Function GetCurrentDate() As Date
    GetCurrentDate = Date
End Function

Function GetCurrentDateTime() As String
    GetCurrentDateTime = Format(Now(), "yyyy-mm-dd hh:mm:ss")
End Function

Function CalculateDaysBetween(startDate As Date, endDate As Date) As Long
    CalculateDaysBetween = DateDiff("d", startDate, endDate)
End Function

Function CalculateWeeksBetween(startDate As Date, endDate As Date) As Long
    CalculateWeeksBetween = DateDiff("w", startDate, endDate)
End Function

Function AddDays(baseDate As Date, daysToAdd As Long) As Date
    AddDays = DateAdd("d", daysToAdd, baseDate)
End Function

Function GetMonthName(monthNumber As Integer) As String
    GetMonthName = Format(DateSerial(2020, monthNumber, 1), "MMMM")
End Function

' ========== FILE OPERATIONS ==========

Function FileExists(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function

Function FolderExists(folderPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
End Function

Sub CreateFolder(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder (folderPath)
    End If
End Sub

Sub DeleteFile(filePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath
    End If
End Sub

Function GetFileSize(filePath As String) As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        GetFileSize = fso.GetFile(filePath).Size
    Else
        GetFileSize = 0
    End If
End Function

' ========== WORKSHEET UTILITIES ==========

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Sub CreateSheetIfNotExists(sheetName As String)
    If Not SheetExists(sheetName) Then
        ThisWorkbook.Sheets.Add.Name = sheetName
    End If
End Sub

Sub DeleteSheet(sheetName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Sub ClearSheetData(ws As Worksheet)
    ws.Cells.Clear
End Sub

Sub AutoFitColumns(ws As Worksheet, Optional lastColumn As String = "Z")
    ws.Columns("A:" & lastColumn).AutoFit
End Sub

Function GetLastRow(ws As Worksheet, Optional columnNumber As Long = 1) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
End Function

Function GetLastColumn(ws As Worksheet, Optional rowNumber As Long = 1) As Long
    GetLastColumn = ws.Cells(rowNumber, ws.Columns.Count).End(xlToLeft).Column
End Function

' ========== RANGE UTILITIES ==========

Sub HighlightRange(rng As Range, colorRGB As Long)
    rng.Interior.Color = colorRGB
End Sub

Sub BoldRange(rng As Range)
    rng.Font.Bold = True
End Sub

Sub SetRangeFontSize(rng As Range, fontSize As Long)
    rng.Font.Size = fontSize
End Sub

Sub MergeRangeCells(rng As Range)
    rng.Merge
End Sub

Sub CenterAlignRange(rng As Range)
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
End Sub

Sub AddBorderToRange(rng As Range)
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Weight = xlMedium
End Sub

Function GetCellValue(ws As Worksheet, row As Long, column As Long) As Variant
    GetCellValue = ws.Cells(row, column).Value
End Function

Sub SetCellValue(ws As Worksheet, row As Long, column As Long, value As Variant)
    ws.Cells(row, column).Value = value
End Sub

' ========== VALIDATION UTILITIES ==========

Function IsValidEmail(email As String) As Boolean
    Dim atPos As Long
    Dim dotPos As Long
    
    atPos = InStr(email, "@")
    dotPos = InStr(atPos, email, ".")
    
    If atPos > 0 And dotPos > atPos + 1 Then
        IsValidEmail = True
    Else
        IsValidEmail = False
    End If
End Function

Function IsValidNumber(value As Variant) As Boolean
    On Error Resume Next
    IsValidNumber = Not IsNull(CDbl(value))
    On Error GoTo 0
End Function

Function IsValidPositiveNumber(value As Variant) As Boolean
    If IsValidNumber(value) Then
        IsValidPositiveNumber = CDbl(value) > 0
    Else
        IsValidPositiveNumber = False
    End If
End Function

Function IsEmptyCell(ws As Worksheet, row As Long, column As Long) As Boolean
    IsEmptyCell = IsEmpty(ws.Cells(row, column).Value)
End Function

' ========== DATA EXPORT UTILITIES ==========

Sub ExportRangeToCSV(rng As Range, filePath As String)
    Dim fso As Object
    Dim textFile As Object
    Dim row As Range
    Dim cell As Range
    Dim lineData As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set textFile = fso.CreateTextFile(filePath, True)
    
    For Each row In rng.Rows
        lineData = ""
        For Each cell In row.Cells
            lineData = lineData & cell.Value & ","
        Next cell
        textFile.WriteLine Left(lineData, Len(lineData) - 1)
    Next row
    
    textFile.Close
End Sub

Sub ExportRangeToText(rng As Range, filePath As String)
    Dim fso As Object
    Dim textFile As Object
    Dim row As Range
    Dim cell As Range
    Dim lineData As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set textFile = fso.CreateTextFile(filePath, True)
    
    For Each row In rng.Rows
        lineData = ""
        For Each cell In row.Cells
            lineData = lineData & cell.Value & vbTab
        Next cell
        textFile.WriteLine Trim(lineData)
    Next row
    
    textFile.Close
End Sub

' ========== ERROR HANDLING UTILITIES ==========

Sub LogError(errorNumber As Long, errorDescription As String, functionName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ErrorLog")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "ErrorLog"
        ws.Range("A1").Value = "Timestamp"
        ws.Range("B1").Value = "Function"
        ws.Range("C1").Value = "Error Number"
        ws.Range("D1").Value = "Error Description"
    End If
    On Error GoTo 0
    
    lastRow = GetLastRow(ws) + 1
    
    ws.Range("A" & lastRow).Value = Now()
    ws.Range("B" & lastRow).Value = functionName
    ws.Range("C" & lastRow).Value = errorNumber
    ws.Range("D" & lastRow).Value = errorDescription
End Sub

Function GetErrorDescription(errorNumber As Long) As String
    Select Case errorNumber
        Case 9
            GetErrorDescription = "Subscript out of range"
        Case 11
            GetErrorDescription = "Division by zero"
        Case 13
            GetErrorDescription = "Type mismatch"
        Case 424
            GetErrorDescription = "Object required"
        Case Else
            GetErrorDescription = "Error #" & errorNumber
    End Select
End Function

' ========== SYSTEM UTILITIES ==========

Sub SetCalculationMode(automaticMode As Boolean)
    If automaticMode Then
        Application.Calculation = xlAutomatic
    Else
        Application.Calculation = xlManual
    End If
End Sub

Sub DisableScreenUpdates()
    Application.ScreenUpdating = False
End Sub

Sub EnableScreenUpdates()
    Application.ScreenUpdating = True
End Sub

Sub ShowStatusMessage(message As String)
    Application.StatusBar = message
End Sub

Sub ClearStatusBar()
    Application.StatusBar = ""
End Sub

' ========== WORKBOOK UTILITIES ==========

Sub SaveWorkbook()
    ThisWorkbook.Save
    MsgBox "Workbook saved successfully!", vbInformation
End Sub

Sub SaveWorkbookAs(filePath As String)
    ThisWorkbook.SaveAs filePath
    MsgBox "Workbook saved as: " & filePath, vbInformation
End Sub

Function GetWorkbookPath() As String
    GetWorkbookPath = ThisWorkbook.Path
End Function

Function GetWorkbookName() As String
    GetWorkbookName = ThisWorkbook.Name
End Function

Sub CloseWorkbook()
    ThisWorkbook.Close
End Sub

' ========== BACKUP AND RECOVERY ==========

Sub CreateWorkbookBackup()
    Dim backupPath As String
    Dim fso As Object
    Dim sourcePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    sourcePath = ThisWorkbook.FullName
    backupPath = GetWorkbookPath() & "\Backup_" & Format(Now(), "yyyymmdd_hhmmss") & ".xlsx"
    
    fso.CopyFile sourcePath, backupPath
    
    MsgBox "Backup created at: " & backupPath, vbInformation
End Sub

Sub RestoreFromBackup(backupPath As String)
    Dim fso As Object
    Dim currentPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    currentPath = ThisWorkbook.FullName
    
    If fso.FileExists(backupPath) Then
        fso.DeleteFile currentPath
        fso.CopyFile backupPath, currentPath
        MsgBox "Restored from backup: " & backupPath, vbInformation
    Else
        MsgBox "Backup file not found: " & backupPath, vbExclamation
    End If
End Sub

' ========== QUICK OPERATIONS ==========

Sub OpenExplorer()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Shell "explorer.exe /e," & GetWorkbookPath()
End Sub

Sub CopyToClipboard(text As String)
    Dim dataObj As Object
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    dataObj.SetText text
    dataObj.PutInClipboard
End Sub

Function GetFromClipboard() As String
    Dim dataObj As Object
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    GetFromClipboard = dataObj.GetFromClipboard()
End Function

Sub DebugPrint(message As String)
    Debug.Print "[" & Format(Now(), "hh:mm:ss") & "] " & message
End Sub

Sub ShowDebugInfo()
    Dim infoText As String
    
    infoText = "Workbook: " & GetWorkbookName() & vbCrLf & _
               "Path: " & GetWorkbookPath() & vbCrLf & _
               "Sheets: " & ThisWorkbook.Sheets.Count & vbCrLf & _
               "Last Saved: " & Format(ThisWorkbook.BuiltinDocumentProperties("Last Save Time").Value, "yyyy-mm-dd hh:mm:ss")
    
    MsgBox infoText, vbInformation, "Debug Information"
End Sub
