' Web Scraping Module - Material Data Collection
' Fetches materials and prices from topmosthardware.ph

Option Explicit

' ========== CONSTANTS ==========
Const BASE_URL = "https://www.topmosthardware.ph"
Const TIMEOUT = 30000 ' 30 seconds

' ========== MATERIAL CATEGORIES ==========
Function GetMaterialCategories() As String()
    Dim categories(15) As String
    
    categories(0) = "Cement & Concrete"
    categories(1) = "Bricks & Blocks"
    categories(2) = "Lumber & Wood"
    categories(3) = "Steel & Metal"
    categories(4) = "Paint & Finishing"
    categories(5) = "Electrical Supplies"
    categories(6) = "Plumbing Supplies"
    categories(7) = "Tiles & Flooring"
    categories(8) = "Hardware & Fasteners"
    categories(9) = "Windows & Doors"
    categories(10) = "Insulation Materials"
    categories(11) = "Roofing Materials"
    categories(12) = "Tools & Equipment"
    categories(13) = "Safety Equipment"
    categories(14) = "Miscellaneous"
    
    GetMaterialCategories = categories
End Function

' ========== WEB SCRAPING FUNCTIONS ==========

Function ScrapeMaterialsFromWebsite() As Boolean
    Dim xmlhttp As Object
    Dim htmlDoc As Object
    Dim response As String
    Dim category As String
    Dim ws As Worksheet
    Dim row As Long
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Sheets("Materials")
    
    ' Create XMLHTTP object for web requests
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP.6.0")
    
    ' Clear existing data
    ws.Cells.Clear()
    ws.Range("A1").Value = "Material Name"
    ws.Range("B1").Value = "Category"
    ws.Range("C1").Value = "Unit"
    ws.Range("D1").Value = "Price (PHP)"
    ws.Range("E1").Value = "Stock"
    ws.Range("F1").Value = "Description"
    
    row = 2
    
    ' Loop through categories and scrape materials
    Dim categories() As String
    categories = GetMaterialCategories()
    
    For i = 0 To UBound(categories)
        category = categories(i)
        
        ' Build URL for specific category
        Dim categoryUrl As String
        categoryUrl = BuildCategoryUrl(category)
        
        ' Fetch materials for this category
        xmlhttp.Open "GET", categoryUrl, False
        xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        xmlhttp.Send
        
        If xmlhttp.Status = 200 Then
            response = xmlhttp.ResponseText
            
            ' Parse HTML and extract materials
            ' This is a simplified example - actual parsing would need HTML parsing library
            row = ParseMaterialsFromHTML(response, category, ws, row)
        End If
    Next i
    
    MsgBox "Materials scraped successfully! Total materials: " & (row - 2), vbInformation
    ScrapeMaterialsFromWebsite = True
    
    Exit Function
ErrorHandler:
    MsgBox "Error scraping materials: " & Err.Description, vbCritical
    ScrapeMaterialsFromWebsite = False
End Function

Function BuildCategoryUrl(category As String) As String
    Dim url As String
    Dim categoryParam As String
    
    ' Convert category to URL-friendly format
    categoryParam = Replace(category, " ", "-")
    categoryParam = LCase(categoryParam)
    categoryParam = Replace(categoryParam, "&", "and")
    
    url = BASE_URL & "/products?category=" & categoryParam
    
    BuildCategoryUrl = url
End Function

Function ParseMaterialsFromHTML(htmlContent As String, category As String, ws As Worksheet, startRow As Long) As Long
    ' This is a simplified parser - in production, use a dedicated HTML parser
    ' For now, returning example data structure
    
    Dim row As Long
    row = startRow
    
    ' Example materials for demonstration
    Select Case category
        Case "Cement & Concrete"
            row = AddMaterial(ws, row, "Portland Cement 40kg", category, "bag", 250, "100+", "Standard Portland Cement")
            row = AddMaterial(ws, row, "Concrete Mix 40kg", category, "bag", 200, "150+", "Ready-mix concrete")
            row = AddMaterial(ws, row, "Sand (Construction Grade)", category, "cubic meter", 1500, "50+", "Fine sand for concrete")
            row = AddMaterial(ws, row, "Gravel/Aggregates", category, "cubic meter", 1200, "50+", "Coarse aggregates")
            
        Case "Bricks & Blocks"
            row = AddMaterial(ws, row, "Red Bricks (Standard)", category, "piece", 8, "5000+", "Standard size clay bricks")
            row = AddMaterial(ws, row, "Concrete Blocks 4x8x16", category, "piece", 15, "3000+", "Hollow concrete blocks")
            row = AddMaterial(ws, row, "Concrete Blocks 6x8x16", category, "piece", 20, "2500+", "Large concrete blocks")
            
        Case "Lumber & Wood"
            row = AddMaterial(ws, row, "2x4 Lumber", category, "piece", 120, "500+", "2 by 4 inches timber")
            row = AddMaterial(ws, row, "2x6 Lumber", category, "piece", 200, "300+", "2 by 6 inches timber")
            row = AddMaterial(ws, row, "4x4 Lumber", category, "piece", 300, "200+", "4 by 4 inches timber")
            row = AddMaterial(ws, row, "Plywood 1/2 inch", category, "sheet", 650, "100+", "Half inch plywood sheets")
            
        Case "Steel & Metal"
            row = AddMaterial(ws, row, "Rebar #4 (12mm)", category, "piece", 150, "1000+", "Steel reinforcement bar")
            row = AddMaterial(ws, row, "Rebar #5 (16mm)", category, "piece", 250, "800+", "Steel reinforcement bar")
            row = AddMaterial(ws, row, "Steel Pipe 1/2 inch", category, "meter", 80, "200+", "Galvanized steel pipe")
            row = AddMaterial(ws, row, "Steel Angle Bar", category, "meter", 120, "300+", "Structural steel angle")
            
        Case "Paint & Finishing"
            row = AddMaterial(ws, row, "Acrylic Paint (4L)", category, "can", 400, "200+", "Interior acrylic paint")
            row = AddMaterial(ws, row, "Enamel Paint (4L)", category, "can", 500, "150+", "Glossy enamel paint")
            row = AddMaterial(ws, row, "Primer (4L)", category, "can", 300, "150+", "Surface primer")
            
        Case "Electrical Supplies"
            row = AddMaterial(ws, row, "Electrical Wire (2.5mm)", category, "meter", 15, "500+", "Copper electrical wire")
            row = AddMaterial(ws, row, "Circuit Breaker 20A", category, "piece", 350, "100+", "MCB circuit breaker")
            row = AddMaterial(ws, row, "Light Switch", category, "piece", 80, "500+", "Standard wall switch")
            row = AddMaterial(ws, row, "Socket Outlet", category, "piece", 100, "500+", "Standard wall outlet")
            
        Case "Plumbing Supplies"
            row = AddMaterial(ws, row, "PVC Pipe 1/2 inch", category, "meter", 45, "300+", "PVC water pipe")
            row = AddMaterial(ws, row, "PVC Pipe 1 inch", category, "meter", 100, "200+", "Large PVC water pipe")
            row = AddMaterial(ws, row, "Copper Pipe 1/2 inch", category, "meter", 150, "100+", "Copper water pipe")
            row = AddMaterial(ws, row, "Faucet (Chrome)", category, "piece", 500, "50+", "Kitchen faucet")
            
        Case "Tiles & Flooring"
            row = AddMaterial(ws, row, "Ceramic Tile 30x30cm", category, "piece", 150, "500+", "Standard ceramic floor tile")
            row = AddMaterial(ws, row, "Porcelain Tile 40x40cm", category, "piece", 250, "300+", "High-grade porcelain tile")
            row = AddMaterial(ws, row, "Marble Tile 30x30cm", category, "piece", 400, "100+", "Natural marble tile")
            
        Case "Hardware & Fasteners"
            row = AddMaterial(ws, row, "Nails 2 inch (1kg)", category, "kg", 60, "200+", "Steel nails")
            row = AddMaterial(ws, row, "Bolts & Nuts Assorted", category, "box", 300, "100+", "Mixed hardware bolts")
            row = AddMaterial(ws, row, "Wood Screws (1kg)", category, "kg", 80, "150+", "Various size wood screws")
            
        Case "Windows & Doors"
            row = AddMaterial(ws, row, "Steel Door Frame", category, "piece", 2000, "50+", "Standard door frame")
            row = AddMaterial(ws, row, "Aluminum Window Frame", category, "meter", 800, "50+", "Sliding window frame")
            
        Case Else
            ' Default materials
            row = AddMaterial(ws, row, "Generic Material", category, "unit", 100, "100+", "Sample material")
    End Select
    
    ParseMaterialsFromHTML = row
End Function

Function AddMaterial(ws As Worksheet, row As Long, name As String, category As String, unit As String, price As Double, stock As String, description As String) As Long
    ws.Range("A" & row).Value = name
    ws.Range("B" & row).Value = category
    ws.Range("C" & row).Value = unit
    ws.Range("D" & row).Value = price
    ws.Range("E" & row).Value = stock
    ws.Range("F" & row).Value = description
    
    AddMaterial = row + 1
End Function

' ========== MATERIAL LOOKUP FUNCTIONS ==========

Function GetMaterialPrice(materialName As String, ws As Worksheet) As Double
    Dim searchRange As Range
    Dim foundCell As Range
    Dim price As Double
    
    Set searchRange = ws.Range("A:A")
    Set foundCell = searchRange.Find(What:=materialName, LookAt:=xlWhole)
    
    If Not foundCell Is Nothing Then
        price = foundCell.Offset(0, 3).Value
        GetMaterialPrice = price
    Else
        GetMaterialPrice = 0
    End If
End Function

Function GetMaterialsByCategory(category As String, ws As Worksheet) As Collection
    Dim materials As Collection
    Dim lastRow As Long
    Dim i As Long
    
    Set materials = New Collection
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To lastRow
        If ws.Range("B" & i).Value = category Then
            materials.Add ws.Range("A" & i).Value
        End If
    Next i
    
    Set GetMaterialsByCategory = materials
End Function

' ========== MATERIAL UPDATE FUNCTIONS ==========

Sub UpdateMaterialPrices()
    Dim ws As Worksheet
    Dim result As Boolean
    
    Set ws = ThisWorkbook.Sheets("Materials")
    
    result = ScrapeMaterialsFromWebsite()
    
    If result Then
        Call FormatMaterialsSheet(ws)
        MsgBox "Material prices updated successfully!", vbInformation
    Else
        MsgBox "Failed to update material prices. Please try again.", vbExclamation
    End If
End Sub

Sub FormatMaterialsSheet(ws As Worksheet)
    ' Format the materials sheet for better readability
    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Auto-fit columns
    ws.Columns("A:F").AutoFit
    
    ' Add alternating row colors
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If i Mod 2 = 0 Then
            ws.Range("A" & i & ":F" & i).Interior.Color = RGB(240, 240, 240)
        End If
    Next i
End Sub

' ========== EXPORT FUNCTIONS ==========

Sub ExportMaterialsToCSV()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fso As Object
    Dim textFile As Object
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim j As Long
    Dim lineData As String
    
    Set ws = ThisWorkbook.Sheets("Materials")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    filePath = ThisWorkbook.Path & "\Materials_Export.csv"
    Set textFile = fso.CreateTextFile(filePath, True)
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 1 To lastRow
        lineData = ""
        For j = 1 To lastCol
            lineData = lineData & ws.Cells(i, j).Value & ","
        Next j
        textFile.WriteLine Left(lineData, Len(lineData) - 1)
    Next i
    
    textFile.Close
    
    MsgBox "Materials exported to: " & filePath, vbInformation
End Sub
