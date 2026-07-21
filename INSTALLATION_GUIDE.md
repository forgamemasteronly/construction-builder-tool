# Construction Builder Tool - Installation & Setup Guide

## Overview
The Construction Builder Tool is a comprehensive Excel-based VBA application designed to calculate construction project costs, labor expenses, and material requirements. It supports both Philippine Peso (PHP) and US Dollar (USD) currencies with automatic exchange rate conversion.

---

## Table of Contents
1. [System Requirements](#system-requirements)
2. [Installation Instructions](#installation-instructions)
3. [Initial Setup](#initial-setup)
4. [Module Overview](#module-overview)
5. [How to Use](#how-to-use)
6. [Features](#features)
7. [Troubleshooting](#troubleshooting)
8. [FAQ](#faq)

---

## System Requirements

### Hardware
- **Processor**: Intel i5 or equivalent (or better)
- **RAM**: Minimum 4GB (8GB recommended)
- **Storage**: 50MB free disk space
- **Display**: 1366x768 resolution minimum

### Software
- **Operating System**: Windows 7 or later (Windows 10/11 recommended)
- **Microsoft Office**: Excel 2016 or later (Office 365 recommended)
- **Internet Connection**: Required for live exchange rate updates

### Excel Configuration
1. Enable macros in Excel
2. Ensure XMLHTTP is available (usually enabled by default)
3. Scripting Runtime library must be enabled

---

## Installation Instructions

### Step 1: Download the Tool
```
1. Go to: https://github.com/forgamemasteronly/Construction-Builder-Tool
2. Click "Code" → "Download ZIP"
3. Extract the ZIP file to your desired location
```

### Step 2: Prepare the Excel Workbook
```
1. Open Microsoft Excel
2. Create a new workbook or use an existing one
3. Save it as "Construction_Builder.xlsm" (note: must be .xlsm format)
```

### Step 3: Import VBA Modules
```
1. Press Alt + F11 to open VBA Editor
2. Go to File → Import File
3. Navigate to the VBA_Modules folder
4. Import the following modules in order:
   a) Utilities.bas
   b) ConstructionBuilder_Main.bas
   c) ExchangeRate.bas
   d) LaborCost.bas
   e) WebScraping.bas
   f) UserInterface.bas
```

### Step 4: Enable Trust Center Settings
```
1. Go to File → Options → Trust Center
2. Click "Trust Center Settings"
3. Go to Macro Settings
4. Select "Enable all macros"
5. Check "Trust access to the VBA object model"
6. Click OK
```

### Step 5: Verify Installation
```
1. Close VBA Editor
2. Press Alt + F8 to view available macros
3. You should see all module functions listed
4. Save the file as .xlsm
```

---

## Initial Setup

### First Run Configuration

#### Step 1: Initialize Dashboard
```
1. In Excel, press Alt + F8
2. Select "InitializeDashboard"
3. Click Run
4. This creates the main dashboard sheet
```

#### Step 2: Configure Project Settings
```
1. Go to the Dashboard sheet
2. Enter your Project Name in cell B2
3. Select Country (Philippines or USA) in cell B5
4. The currency will auto-select based on country
5. Check Exchange Rate in cell B7
```

#### Step 3: Load Material Database
```
1. Press Alt + F8
2. Select "UpdateMaterialPrices"
3. Click Run
4. Materials will be fetched from topmosthardware.ph
5. Check the Materials sheet for the complete list
```

---

## Module Overview

### 1. **Utilities.bas** - Helper Functions
**Purpose**: Provides reusable utility functions across all modules

**Key Functions**:
- String manipulation (TrimString, ToUpperCase, ToLowerCase)
- Number formatting (RoundNumber, FormatNumberToCurrency)
- File operations (FileExists, CreateFolder, DeleteFile)
- Worksheet utilities (SheetExists, GetLastRow, AutoFitColumns)
- Data validation (IsValidEmail, IsValidNumber, IsValidPositiveNumber)
- Clipboard operations
- Error logging

**Usage Example**:
```vba
Dim formatted As String
formatted = FormatNumberToCurrency(15000, "PHP")
' Output: ₱15,000.00
```

---

### 2. **ConstructionBuilder_Main.bas** - Core Calculations
**Purpose**: Main calculation engine for materials and labor

**Key Functions**:
- `CalculateTotalMaterialCost()` - Calculates total material expenses
- `CalculateMaterialCost()` - Individual material cost with waste factor
- `CalculateTotalLaborCost()` - Calculates total labor expenses
- `CalculateProjectCost()` - Total project cost (materials + labor)
- `ConvertCurrency()` - Currency conversion

**Key Constants**:
- `DEFAULT_WASTE_FACTOR = 0.12` (12% waste allowance)
- `PHP_MIN_WAGE = 570` (PHP daily minimum)
- `USD_MIN_WAGE = 96` (USD daily minimum)

**Usage Example**:
```vba
Dim projectCost As Double
projectCost = CalculateProjectCost(materials, workers, 10, "Philippines", 56.5)
MsgBox "Total Project Cost: ₱" & Format(projectCost, "#,##0.00")
```

---

### 3. **ExchangeRate.bas** - Currency Management
**Purpose**: Handles exchange rates and currency conversions

**Key Functions**:
- `GetCurrentExchangeRate()` - Fetches live exchange rates from API
- `ConvertPHPtoUSD()` - PHP to USD conversion
- `ConvertUSDtoPHP()` - USD to PHP conversion
- `GetCachedExchangeRate()` - Retrieves cached rates (1-hour cache)
- `GetAverageExchangeRate()` - Historical average rates

**API Integration**:
- Uses exchangerate-api.com for live rates
- Caches rates for 1 hour to reduce API calls
- Falls back to manual rates if API is unavailable

**Usage Example**:
```vba
Dim rate As Double
rate = GetCurrentExchangeRate("USD", "PHP")
MsgBox "1 USD = ₱" & Format(rate, "0.00")
```

---

### 4. **LaborCost.bas** - Worker Management
**Purpose**: Handles worker classifications and labor calculations

**Key Functions**:
- `GetWorkerClassifications()` - Returns 10 worker types
- `CalculateWorkerTotalCost()` - Total cost for worker type
- `CalculateTotalWithOvertime()` - Includes overtime calculations
- `ApplyPhilippineLaborLaw()` - Adds 13th month pay
- `ApplyUSALaborLaw()` - US overtime rules (1.5x)
- `GetDailyWorkerRate()` - Rate by specialization

**Supported Worker Types**:
1. Armalador (Steel Binder)
2. Albañil (Mason)
3. Tindero (Carpenter)
4. Electricista (Electrician)
5. Plomero (Plumber)
6. Magamit (Equipment Operator)
7. Manggagawa (General Laborer)
8. Pintero (Painter)
9. Weldes (Welder)
10. Tile Layer

**Usage Example**:
```vba
Dim laborCost As Double
laborCost = CalculateWorkerTotalCost("Albañil", 3, 15, "Philippines")
' 3 masons, 15 days, in Philippines
```

---

### 5. **WebScraping.bas** - Material Database
**Purpose**: Fetches material prices from online hardware suppliers

**Key Functions**:
- `ScrapeMaterialsFromWebsite()` - Fetches from topmosthardware.ph
- `GetMaterialsByCategory()` - Filters materials by type
- `GetMaterialPrice()` - Lookup price by material name
- `ExportMaterialsToCSV()` - Export material list

**Material Categories**:
- Cement & Concrete
- Bricks & Blocks
- Lumber & Wood
- Steel & Metal
- Paint & Finishing
- Electrical Supplies
- Plumbing Supplies
- Tiles & Flooring
- Hardware & Fasteners
- Windows & Doors
- And more...

**Usage Example**:
```vba
Sub RefreshMaterials()
    If ScrapeMaterialsFromWebsite() Then
        Call FormatMaterialsSheet(ThisWorkbook.Sheets("Materials"))
    End If
End Sub
```

---

### 6. **UserInterface.bas** - Dashboard & Forms
**Purpose**: Provides user interface and input forms

**Key Functions**:
- `InitializeDashboard()` - Creates main dashboard
- `ShowMaterialInputForm()` - Add materials dialog
- `ShowWorkerInputForm()` - Add workers dialog
- `GenerateProjectReport()` - Creates detailed report
- `ExportToCSV()` - Exports project data
- `RecalculateProjectCost()` - Recalculates all values

**Dashboard Sections**:
1. **Project Settings** - Name, country, currency, exchange rate
2. **Dimensions** - Length, width, height, calculated area/volume
3. **Materials** - List with quantities, prices, and totals
4. **Labor** - Worker types, counts, days, and costs
5. **Summary** - Total costs with 10% contingency

**Usage Example**:
```vba
Sub Main()
    Call InitializeDashboard
    Call UpdateMaterialPrices
    Call RecalculateProjectCost
End Sub
```

---

## How to Use

### Basic Workflow

#### 1. Create a New Project
```
1. Open the Construction_Builder.xlsm file
2. Go to Dashboard sheet
3. Enter Project Name
4. Select Country (Philippines or USA)
5. Enter Dimensions (Length, Width, Height, Thickness)
```

#### 2. Add Materials
```
1. Click "Add Material" button (or use macro)
2. Enter:
   - Material Name (e.g., "Portland Cement 40kg")
   - Quantity needed
   - Unit Price
3. Materials automatically calculate costs with 12% waste factor
```

#### 3. Add Workers
```
1. Click "Add Worker" button
2. Enter:
   - Worker Specialization (e.g., "Albañil")
   - Number of Workers
   - Days Required
3. System calculates labor costs with applicable laws
```

#### 4. Review Summary
```
1. Check Project Summary section
2. View:
   - Total Materials Cost
   - Total Labor Cost
   - 10% Contingency
   - TOTAL PROJECT COST
```

#### 5. Generate Report
```
1. Click "Generate Project Report"
2. Report sheet is created with full details
3. Can be printed or exported to PDF
```

#### 6. Export Data
```
1. Click "Export to CSV"
2. File is saved with timestamp
3. Use in other applications (accounting, email, etc.)
```

---

## Features

### ✓ Material Management
- Database of 100+ construction materials
- Automatic web scraping from hardware suppliers
- Price updates with exchange rate conversion
- 12% waste factor included in calculations
- Export/import capabilities

### ✓ Labor Calculations
- 10 professional worker classifications
- Philippines: Includes 13th month pay + legal benefits
- USA: Overtime calculation (1.5x after 40 hours/week)
- Minimum wage enforcement
- Cost comparison PHP vs USD

### ✓ Currency Management
- Live exchange rate fetching from external API
- 1-hour caching to reduce API calls
- Manual rate override option
- Support for PHP and USD
- Historical rate tracking

### ✓ Project Costing
- Automatic material cost calculations
- Labor cost with country-specific laws
- Area and volume calculations
- 10% contingency add-on
- Real-time updates

### ✓ Reporting
- Detailed project reports
- CSV export for spreadsheet analysis
- Print-ready dashboards
- Cost breakdowns
- Date and timestamp tracking

### ✓ Data Validation
- Minimum wage enforcement
- Input validation
- Error logging
- Backup and recovery options

---

## Troubleshooting

### Issue: Macros are disabled
**Solution**:
```
1. File → Options → Trust Center
2. Enable macros
3. Trust access to VBA object model
4. Restart Excel
```

### Issue: "Object required" error
**Solution**:
```
1. Ensure all modules are imported
2. Check Tools → References for missing libraries
3. Enable "Microsoft Scripting Runtime"
4. Restart Excel
```

### Issue: Exchange rate won't update
**Solution**:
```
1. Check internet connection
2. Open Task Manager → verify network access
3. Try manual exchange rate: SetManualExchangeRate()
4. Check firewall settings
```

### Issue: Web scraping returns no materials
**Solution**:
```
1. Verify website is still: topmosthardware.ph
2. Check internet connection
3. Website may require updates
4. Use manual material entry
```

### Issue: Calculations are wrong
**Solution**:
```
1. Verify all inputs are correct numbers
2. Check exchange rate is updated
3. Ensure correct country is selected
4. Recalculate: RecalculateProjectCost()
```

---

## FAQ

### Q: Can I use this for projects outside Philippines/USA?
**A**: Currently optimized for Philippines and USA labor laws. You can modify the LaborCost.bas module to add more countries.

### Q: What if the website is down?
**A**: The tool falls back to manual material entry. You can still use the calculator with manually entered material prices.

### Q: Can I export the project to accounting software?
**A**: Yes! Export to CSV and import into Excel, QuickBooks, or other accounting software.

### Q: Is my data backed up?
**A**: Use CreateWorkbookBackup() regularly. Creates a timestamped backup file.

### Q: How accurate are the calculations?
**A**: Calculations are based on current market rates (via topmosthardware.ph) and official labor rates. Always verify with actual suppliers.

### Q: Can multiple people use this simultaneously?
**A**: This is a single-file solution. For collaboration, share the file via cloud storage but ensure only one person edits at a time.

### Q: How do I update the material prices?
**A**: Run UpdateMaterialPrices() to fetch latest prices from the supplier.

### Q: What is the 12% waste factor?
**A**: This is an industry standard allowance for material wastage during construction (cutting, breakage, theft, etc.).

### Q: Can I customize worker rates?
**A**: Yes! Edit the GetWorkerRate() function in LaborCost.bas or modify rates in the Workers sheet.

### Q: How do I create a backup?
**A**: Run CreateWorkbookBackup() in the Utilities module. Backups are saved in the same folder as the workbook.

---

## Support & Updates

### Getting Help
- Check this guide thoroughly
- Review the FAQ section
- Check error logs in the "ErrorLog" sheet
- Verify all modules are properly imported

### Reporting Issues
- Document the exact steps to reproduce
- Note the error message
- Include your Excel version
- Export your project data for troubleshooting

### Updates
- Check GitHub repository regularly for updates
- Import new module versions when available
- Always backup before updating

---

## Version History

**Version 1.0** (July 2026)
- Initial release
- 6 main modules
- Support for PHP and USD
- Material database integration
- Labor law compliance for Philippines and USA
- Web scraping capabilities
- Comprehensive reporting

---

## License

This tool is provided for construction estimation purposes. Always verify calculations with actual suppliers and labor rates.

---

**Last Updated**: July 2026
**Maintainer**: forgamemasteronly
**Repository**: https://github.com/forgamemasteronly/Construction-Builder-Tool
