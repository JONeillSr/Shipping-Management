<#
.SYNOPSIS
    Excel to CSV Export Helper for Logistics Email Automation (Enhanced)
    
.DESCRIPTION
    This helper script automates the export of auction data from Excel files to CSV format,
    preparing it for use with the Logistics Email Generation script. It handles multiple
    sheets, data validation, filtering, and format conversion with intelligent defaults.
    Enhanced with improved error handling and diagnostic capabilities.
    
.PARAMETER ExcelPath
    Path to the Excel file (.xlsx, .xlsm, .xls)
    
.PARAMETER OutputDirectory
    Directory where CSV files will be saved (default: current directory)
    
.PARAMETER SheetName
    Specific sheet name to export (optional - will prompt if not specified)
    
.PARAMETER FilterAddress
    Filter data by specific pickup address (optional)
    
.PARAMETER ExcludeZeroQuantity
    Exclude rows where quantity is 0 or empty
    
.PARAMETER AddMissingColumns
    Automatically add any missing required columns with default values
    
.PARAMETER OpenAfterExport
    Open the exported CSV file after creation
    
.PARAMETER ValidateImages
    Check if lot images exist in specified directory
    
.PARAMETER DiagnosticMode
    Run in diagnostic mode to troubleshoot issues
    
.EXAMPLE
    .\Export-ExcelToCSV.ps1 -ExcelPath "PurchaseTracking.xlsm"
    
.EXAMPLE
    .\Export-ExcelToCSV.ps1 -ExcelPath "Auction.xlsx" -SheetName "Heartland 901" -FilterAddress "Sturgis, MI"
    
.EXAMPLE
    .\Export-ExcelToCSV.ps1 -ExcelPath "data.xlsx" -DiagnosticMode
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 10-07-2025
    Version: 1.1.0
    Change Date: 10-07-2025
    Change Purpose: Enhanced error handling and diagnostic capabilities

.CHANGELOG
    1.1.0 - Enhanced error handling and diagnostics
          - Added diagnostic mode for troubleshooting
          - Improved empty sheet handling
          - Multiple import fallback methods
          - Better column validation
          - Enhanced user feedback
    
    1.0.0 - Initial release
          - Excel to CSV conversion
          - Multi-sheet support
          - Data filtering options
          - Image validation
          - Interactive mode

.LINK
    https://github.com/JONeillSr/LogisticsAutomation
    
.COMPONENT
    Requires PowerShell 5.1 or higher
    Requires ImportExcel module
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true, Position=0)]
    [ValidateScript({
        if (Test-Path $_) {
            $extension = [System.IO.Path]::GetExtension($_)
            if ($extension -match '\.(xlsx?|xlsm)$') {
                return $true
            }
            throw "File must be an Excel file (.xlsx, .xls, or .xlsm)"
        }
        throw "File not found: $_"
    })]
    [string]$ExcelPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputDirectory = (Get-Location).Path,
    
    [Parameter(Mandatory=$false)]
    [string]$SheetName,
    
    [Parameter(Mandatory=$false)]
    [string]$FilterAddress,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExcludeZeroQuantity,
    
    [Parameter(Mandatory=$false)]
    [switch]$AddMissingColumns,
    
    [Parameter(Mandatory=$false)]
    [switch]$OpenAfterExport,
    
    [Parameter(Mandatory=$false)]
    [string]$ValidateImages,
    
    [Parameter(Mandatory=$false)]
    [switch]$InteractiveMode,
    
    [Parameter(Mandatory=$false)]
    [switch]$DiagnosticMode,
    
    [Parameter(Mandatory=$false)]
    [switch]$ForceRawImport
)

#region Module Management
Write-Host "`n=== Excel to CSV Export Helper v1.1.0 ===" -ForegroundColor Cyan
Write-Host "Enhanced Edition - Azure Innovators`n" -ForegroundColor Gray

# Check and install ImportExcel module if needed
if (!(Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "📦 Installing ImportExcel module..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
        Write-Host "✅ ImportExcel module installed successfully" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to install ImportExcel module. Please run as Administrator or install manually."
        exit 1
    }
}
Import-Module ImportExcel -Force
#endregion

#region Global Variables
$script:Statistics = @{
    TotalRows = 0
    ExportedRows = 0
    FilteredRows = 0
    ColumnsAdded = @()
    MissingImages = @()
    Warnings = @()
    StartTime = Get-Date
    DiagnosticInfo = @{}
}

$script:RequiredColumns = @('Lot', 'Description', 'Address')
$script:OptionalColumns = @('Quantity', 'Location', 'Plant', 'Bid', 'Sale Price', 'Premium', 'Tax')
#endregion

#region Display Functions
function Show-Banner {
    <#
    .SYNOPSIS
        Displays a formatted banner
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced formatting
    #>
    param([string]$Text, [string]$Color = "Cyan")
    
    $border = "=" * 60
    Write-Host $border -ForegroundColor $Color
    Write-Host $Text.PadLeft(($Text.Length + 60) / 2).PadRight(60) -ForegroundColor $Color
    Write-Host $border -ForegroundColor $Color
}

function Write-ColorOutput {
    <#
    .SYNOPSIS
        Writes colored output with icons
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced output formatting
    #>
    param(
        [string]$Message,
        [ValidateSet("Success", "Error", "Warning", "Info", "Processing", "Debug")]
        [string]$Type = "Info"
    )
    
    $icon = switch($Type) {
        "Success"    { "✅"; $color = "Green" }
        "Error"      { "❌"; $color = "Red" }
        "Warning"    { "⚠️ "; $color = "Yellow" }
        "Processing" { "⚙️ "; $color = "Cyan" }
        "Debug"      { "🔍"; $color = "Magenta" }
        default      { "ℹ️ "; $color = "White" }
    }
    
    Write-Host "$icon $Message" -ForegroundColor $color
}
#endregion

#region Diagnostic Functions
function Test-ExcelFile {

    param([string]$Path)
    
    Write-ColorOutput "Running diagnostics on Excel file..." -Type Debug
    
    try {
        # File information
        $fileInfo = Get-Item $Path
        Write-Host "  File: $($fileInfo.Name)" -ForegroundColor Gray
        Write-Host "  Size: $([math]::Round($fileInfo.Length / 1MB, 2)) MB" -ForegroundColor Gray
        Write-Host "  Modified: $($fileInfo.LastWriteTime)" -ForegroundColor Gray
        
        # Try to open Excel package
        $excelPackage = Open-ExcelPackage -Path $Path
        Write-ColorOutput "Excel package opened successfully" -Type Success
        
        # Get worksheet information
        $worksheets = $excelPackage.Workbook.Worksheets
        Write-Host "  Worksheets found: $($worksheets.Count)" -ForegroundColor Gray
        
        foreach ($ws in $worksheets) {
            Write-Host "`n  Sheet: '$($ws.Name)'" -ForegroundColor Cyan
            
            # Check dimensions
            if ($null -eq $ws.Dimension) {
                Write-ColorOutput "    WARNING: Sheet appears to be empty (no dimension)" -Type Warning
                $script:Statistics.Warnings += "Sheet '$($ws.Name)' appears empty"
                continue
            }
            
            $dimension = $ws.Dimension
            Write-Host "    Range: $($dimension.Address)" -ForegroundColor Gray
            Write-Host "    Rows: $($dimension.Rows)" -ForegroundColor Gray
            Write-Host "    Columns: $($dimension.Columns)" -ForegroundColor Gray
            
            # Try to read first row (headers)
            if ($dimension.Rows -gt 0) {
                $headers = @()
                for ($col = 1; $col -le $dimension.Columns; $col++) {
                    $cellValue = $ws.Cells[1, $col].Value
                    if ($cellValue) {
                        $headers += $cellValue
                    }
                }
                
                if ($headers.Count -gt 0) {
                    Write-Host "    Headers found: $($headers -join ', ')" -ForegroundColor Green
                }
                else {
                    Write-ColorOutput "    No headers found in first row" -Type Warning
                }
            }
        }
        
        Close-ExcelPackage $excelPackage -NoSave
        
        Write-ColorOutput "Diagnostic scan complete" -Type Success
        return $true
    }
    catch {
        Write-ColorOutput "Diagnostic error: $_" -Type Error
        return $false
    }
}

function Import-ExcelRawData {
    <#
    .SYNOPSIS
        Attempts raw data import with multiple methods
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: New fallback import method
    #>
    param(
        [string]$Path,
        [string]$Sheet
    )
    
    Write-ColorOutput "Attempting alternative import methods..." -Type Processing
    
    # Method 1: Try without headers first
    try {
        Write-ColorOutput "Method 1: Importing without headers..." -Type Debug
        $data = Import-Excel -Path $Path -WorksheetName $Sheet -NoHeader -DataOnly
        
        if ($data -and $data.Count -gt 0) {
            Write-ColorOutput "Raw data imported: $($data.Count) rows" -Type Success
            
            # Try to identify headers
            $firstRow = $data[0]
            $hasHeaders = $false
            
            # Check if first row looks like headers
            $properties = $firstRow.PSObject.Properties.Name
            foreach ($prop in $properties) {
                if ($firstRow.$prop -match '^(Lot|Description|Address|Location|Quantity)') {
                    $hasHeaders = $true
                    break
                }
            }
            
            if ($hasHeaders) {
                Write-ColorOutput "Headers detected in first row" -Type Info
                # Convert to proper format with headers
                $headers = @()
                foreach ($prop in $properties) {
                    $headers += $firstRow.$prop
                }
                
                # Skip first row and rebuild with headers
                $newData = @()
                for ($i = 1; $i -lt $data.Count; $i++) {
                    $row = [PSCustomObject]@{}
                    $j = 0
                    foreach ($prop in $properties) {
                        if ($headers[$j]) {
                            $row | Add-Member -NotePropertyName $headers[$j] -NotePropertyValue $data[$i].$prop
                        }
                        $j++
                    }
                    $newData += $row
                }
                
                return $newData
            }
            else {
                # No headers detected, create generic ones
                Write-ColorOutput "No headers detected, using generic column names" -Type Warning
                return $data
            }
        }
    }
    catch {
        Write-ColorOutput "Method 1 failed: $_" -Type Debug
    }
    
    # Method 2: Try with StartRow parameter
    try {
        Write-ColorOutput "Method 2: Importing with StartRow parameter..." -Type Debug
        $data = Import-Excel -Path $Path -WorksheetName $Sheet -StartRow 1 -DataOnly
        
        if ($data -and $data.Count -gt 0) {
            Write-ColorOutput "Data imported with StartRow: $($data.Count) rows" -Type Success
            return $data
        }
    }
    catch {
        Write-ColorOutput "Method 2 failed: $_" -Type Debug
    }
    
    # Method 3: Direct cell reading
    try {
        Write-ColorOutput "Method 3: Direct cell reading..." -Type Debug
        $excel = Open-ExcelPackage -Path $Path
        $ws = $excel.Workbook.Worksheets[$Sheet]
        
        if ($null -eq $ws) {
            throw "Worksheet '$Sheet' not found"
        }
        
        if ($null -eq $ws.Dimension) {
            throw "Worksheet appears to be empty"
        }
        
        $rows = $ws.Dimension.Rows
        $cols = $ws.Dimension.Columns
        
        Write-Host "  Reading $rows rows x $cols columns" -ForegroundColor Gray
        
        # Read all data
        $allData = @()
        for ($row = 1; $row -le $rows; $row++) {
            $rowData = [PSCustomObject]@{}
            for ($col = 1; $col -le $cols; $col++) {
                $cellValue = $ws.Cells[$row, $col].Value
                $colName = "Column$col"
                
                # If first row, try to use as headers
                if ($row -eq 1 -and $cellValue) {
                    $script:Statistics.DiagnosticInfo["Header_Col$col"] = $cellValue
                }
                
                $rowData | Add-Member -NotePropertyName $colName -NotePropertyValue $cellValue
            }
            $allData += $rowData
        }
        
        Close-ExcelPackage $excel -NoSave
        
        Write-ColorOutput "Direct read complete: $($allData.Count) rows" -Type Success
        return $allData
    }
    catch {
        Write-ColorOutput "Method 3 failed: $_" -Type Error
        throw "All import methods failed. Please check the Excel file structure."
    }
}
#endregion

#region Excel Processing Functions
function Get-ExcelSheets {
    <#
    .SYNOPSIS
        Retrieves all sheet names from Excel file with error handling
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced error handling
    #>
    param([string]$Path)
    
    try {
        Write-ColorOutput "Reading Excel file structure..." -Type Processing
        
        # Method 1: Using Open-ExcelPackage
        try {
            $excelPackage = Open-ExcelPackage -Path $Path
            $sheets = @()
            
            foreach ($ws in $excelPackage.Workbook.Worksheets) {
                $sheetInfo = @{
                    Name = $ws.Name
                    Index = $ws.Index
                    Hidden = $ws.Hidden
                    Empty = $null -eq $ws.Dimension
                }
                
                if ($DiagnosticMode) {
                    Write-Host "  Found sheet: '$($ws.Name)' (Index: $($ws.Index), Hidden: $($ws.Hidden), Empty: $($sheetInfo.Empty))" -ForegroundColor Gray
                }
                
                if (!$ws.Hidden) {
                    $sheets += $ws.Name
                }
            }
            
            Close-ExcelPackage $excelPackage -NoSave
            
            if ($sheets.Count -eq 0) {
                throw "No visible sheets found in Excel file"
            }
            
            return $sheets
        }
        catch {
            Write-ColorOutput "Primary method failed, trying alternative..." -Type Debug
        }
        
        # Method 2: Using Get-ExcelSheetInfo
        try {
            $sheetInfo = Get-ExcelSheetInfo -Path $Path
            $sheets = $sheetInfo | Where-Object { -not $_.Hidden } | Select-Object -ExpandProperty Name
            
            if ($sheets.Count -eq 0) {
                throw "No visible sheets found in Excel file"
            }
            
            return $sheets
        }
        catch {
            throw "Unable to read Excel sheets: $_"
        }
    }
    catch {
        Write-ColorOutput "Error reading Excel file: $_" -Type Error
        throw
    }
}

function Select-Sheet {
    <#
    .SYNOPSIS
        Interactive sheet selection with validation
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced selection interface
    #>
    param([array]$Sheets)
    
    if ($Sheets.Count -eq 0) {
        throw "No sheets available for selection"
    }
    
    Write-Host "`n📊 Available Sheets:" -ForegroundColor Cyan
    Write-Host "────────────────────" -ForegroundColor Gray
    
    for ($i = 0; $i -lt $Sheets.Count; $i++) {
        # Clean sheet name display
        $displayName = $Sheets[$i]
        if ($displayName.Length -gt 50) {
            $displayName = $displayName.Substring(0, 47) + "..."
        }
        Write-Host "  [$($i + 1)] $displayName" -ForegroundColor White
    }
    
    if ($DiagnosticMode) {
        Write-Host "`n  [D] Run diagnostics on all sheets" -ForegroundColor Magenta
    }
    
    Write-Host ""
    do {
        $selection = Read-Host "Select sheet number (1-$($Sheets.Count))$(if($DiagnosticMode){' or D for diagnostics'})"
        
        if ($DiagnosticMode -and $selection -eq 'D') {
            Test-ExcelFile -Path $ExcelPath
            Write-Host ""
            continue
        }
        
        $valid = $selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $Sheets.Count
        if (!$valid) {
            Write-ColorOutput "Invalid selection. Please enter a number between 1 and $($Sheets.Count)" -Type Warning
        }
    } while (!$valid)
    
    return $Sheets[[int]$selection - 1]
}

function Import-ExcelData {
    <#
    .SYNOPSIS
        Imports data from Excel with multiple fallback methods
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced import with fallback methods
    #>
    param(
        [string]$Path,
        [string]$Sheet
    )
    
    try {
        Write-ColorOutput "Importing data from sheet: $Sheet" -Type Processing
        
        # Primary import method
        try {
            $data = Import-Excel -Path $Path -WorksheetName $Sheet -DataOnly
            
            if (!$data) {
                throw "No data returned from sheet"
            }
            
            if ($data.Count -eq 0) {
                throw "Sheet appears to be empty"
            }
            
            # Validate data structure
            $firstRow = $data[0]
            if ($null -eq $firstRow) {
                throw "Invalid data structure"
            }
            
            $properties = $firstRow.PSObject.Properties.Name
            if ($properties.Count -eq 0) {
                throw "No columns detected"
            }
            
            $script:Statistics.TotalRows = $data.Count
            Write-ColorOutput "Found $($data.Count) rows with $($properties.Count) columns" -Type Success
            
            # Display sample data
            if ($data.Count -gt 0) {
                Write-Host "`n📋 Sample Data (First 3 rows):" -ForegroundColor Cyan
                Write-Host "────────────────────────────────" -ForegroundColor Gray
                
                $data | Select-Object -First 3 | Format-Table -AutoSize | Out-String | Write-Host
            }
            
            return $data
        }
        catch {
            Write-ColorOutput "Standard import failed: $_" -Type Warning
            
            if ($ForceRawImport -or $DiagnosticMode) {
                Write-ColorOutput "Attempting raw import methods..." -Type Processing
                $data = Import-ExcelRawData -Path $Path -Sheet $Sheet
                
                if ($data) {
                    $script:Statistics.TotalRows = $data.Count
                    Write-ColorOutput "Raw import successful: $($data.Count) rows" -Type Success
                    return $data
                }
            }
            
            throw "Unable to import data from sheet '$Sheet'"
        }
    }
    catch {
        Write-ColorOutput "Error importing Excel data: $_" -Type Error
        
        if (!$DiagnosticMode) {
            Write-Host "`n💡 Tip: Run with -DiagnosticMode flag for detailed analysis" -ForegroundColor Yellow
            Write-Host "   Example: .\Export-ExcelToCSV.ps1 -ExcelPath `"$Path`" -DiagnosticMode" -ForegroundColor Gray
        }
        
        throw
    }
}

function Test-RequiredColumns {
    <#
    .SYNOPSIS
        Validates required columns exist with flexible matching
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Flexible column matching
    #>
    param($Data)
    
    Write-ColorOutput "Validating column structure..." -Type Processing
    
    if (!$Data -or $Data.Count -eq 0) {
        throw "No data to validate"
    }
    
    $columns = $Data[0].PSObject.Properties.Name
    Write-Host "  Columns found: $($columns -join ', ')" -ForegroundColor Gray
    
    $missingRequired = @()
    $missingOptional = @()
    $columnMapping = @{}
    
    # Check for required columns (case-insensitive)
    foreach ($reqCol in $script:RequiredColumns) {
        $found = $false
        foreach ($col in $columns) {
            if ($col -ieq $reqCol) {
                $found = $true
                $columnMapping[$reqCol] = $col
                break
            }
        }
        
        if (!$found) {
            # Try partial match
            $partialMatch = $columns | Where-Object { $_ -like "*$reqCol*" } | Select-Object -First 1
            if ($partialMatch) {
                Write-ColorOutput "Mapped '$partialMatch' to required column '$reqCol'" -Type Info
                $columnMapping[$reqCol] = $partialMatch
                $found = $true
            }
        }
        
        if (!$found) {
            $missingRequired += $reqCol
        }
    }
    
    # Check optional columns
    foreach ($optCol in $script:OptionalColumns) {
        $found = $columns | Where-Object { $_ -ieq $optCol }
        if (!$found) {
            $missingOptional += $optCol
        }
    }
    
    if ($missingRequired.Count -gt 0) {
        Write-ColorOutput "Missing REQUIRED columns: $($missingRequired -join ', ')" -Type Error
        
        if ($AddMissingColumns) {
            Write-ColorOutput "Adding missing required columns with default values..." -Type Processing
            foreach ($col in $missingRequired) {
                Add-MissingColumn -Data $Data -ColumnName $col
            }
        }
        else {
            Write-Host "`n💡 Tip: Use -AddMissingColumns switch to automatically add missing columns" -ForegroundColor Yellow
            
            # Suggest possible column mappings
            Write-Host "`nPossible column mappings:" -ForegroundColor Cyan
            foreach ($reqCol in $missingRequired) {
                $similar = $columns | Where-Object { $_ -like "*$($reqCol.Substring(0, [Math]::Min(3, $reqCol.Length)))*" }
                if ($similar) {
                    Write-Host "  '$reqCol' might be: $($similar -join ', ')" -ForegroundColor Gray
                }
            }
            
            throw "Required columns missing. Cannot proceed without: $($missingRequired -join ', ')"
        }
    }
    
    if ($missingOptional.Count -gt 0 -and $DiagnosticMode) {
        Write-ColorOutput "Missing optional columns: $($missingOptional -join ', ')" -Type Warning
        $script:Statistics.Warnings += "Optional columns not found: $($missingOptional -join ', ')"
    }
    
    Write-ColorOutput "Column validation complete" -Type Success
    return $true
}

function Add-MissingColumn {
    <#
    .SYNOPSIS
        Adds missing column with intelligent default values
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Intelligent default values
    #>
    param(
        $Data,
        [string]$ColumnName
    )
    
    # Intelligent defaults based on column name
    $defaultValue = switch($ColumnName) {
        "Lot" { 
            # Generate sequential lot numbers
            $startNum = 1000 + (Get-Random -Maximum 8999)
            $counter = 0
            foreach ($row in $Data) {
                $row | Add-Member -NotePropertyName $ColumnName -NotePropertyValue ($startNum + $counter) -Force
                $counter++
            }
            Write-ColorOutput "Added column '$ColumnName' with sequential lot numbers starting at $startNum" -Type Success
            $script:Statistics.ColumnsAdded += $ColumnName
            return
        }
        "Description" { 
            $desc = Read-Host "Enter default description for items (or press Enter for 'Item Description Required')"
            if ([string]::IsNullOrWhiteSpace($desc)) {
                $desc = "Item Description Required"
            }
            $desc
        }
        "Address" { 
            Write-Host "Default pickup address needed for column '$ColumnName'" -ForegroundColor Yellow
            $addr = Read-Host "Enter pickup address"
            if ([string]::IsNullOrWhiteSpace($addr)) {
                $addr = "1234 Main Street, City, ST 12345"
            }
            $addr
        }
        "Quantity" { 1 }
        "Location" { "TBD" }
        default { "" }
    }
    
    if ($ColumnName -ne "Lot") {
        foreach ($row in $Data) {
            $row | Add-Member -NotePropertyName $ColumnName -NotePropertyValue $defaultValue -Force
        }
        
        $script:Statistics.ColumnsAdded += $ColumnName
        Write-ColorOutput "Added column '$ColumnName' with default value: $defaultValue" -Type Success
    }
}
#endregion

#region Data Processing Functions
function Apply-DataFilters {
    <#
    .SYNOPSIS
        Applies user-specified filters to data with validation
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced filtering
    #>
    param($Data)
    
    if (!$Data -or $Data.Count -eq 0) {
        Write-ColorOutput "No data to filter" -Type Warning
        return @()
    }
    
    $filteredData = $Data
    $originalCount = $Data.Count
    
    # Filter by address
    if ($FilterAddress) {
        Write-ColorOutput "Filtering by address: $FilterAddress" -Type Processing
        $filteredData = $filteredData | Where-Object { 
            $_.Address -like "*$FilterAddress*" 
        }
        Write-ColorOutput "Rows matching address filter: $($filteredData.Count)" -Type Info
    }
    
    # Exclude zero quantity
    if ($ExcludeZeroQuantity) {
        Write-ColorOutput "Excluding zero/empty quantity items..." -Type Processing
        $filteredData = $filteredData | Where-Object { 
            $null -ne $_.Quantity -and [string]$_.Quantity -ne '' -and [int]$_.Quantity -gt 0 
        }
        Write-ColorOutput "Rows with quantity > 0: $($filteredData.Count)" -Type Info
    }
    
    # Remove empty rows (all properties null or empty)
    $filteredData = $filteredData | Where-Object {
        $hasData = $false
        foreach ($prop in $_.PSObject.Properties) {
            if (![string]::IsNullOrWhiteSpace($prop.Value)) {
                $hasData = $true
                break
            }
        }
        $hasData
    }
    
    $script:Statistics.FilteredRows = $originalCount - $filteredData.Count
    
    if ($filteredData.Count -eq 0) {
        Write-ColorOutput "No data remaining after filters applied!" -Type Warning
        $continue = Read-Host "Continue anyway? (Y/N)"
        if ($continue -ne 'Y') {
            throw "Export cancelled - no data to export"
        }
    }
    
    return $filteredData
}

function Test-LotImages {
    <#
    .SYNOPSIS
        Validates lot images exist
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced image validation
    #>
    param(
        $Data,
        [string]$ImagePath
    )
    
    if (!$ValidateImages -or !$ImagePath) { return }
    
    if (!(Test-Path $ImagePath)) {
        Write-ColorOutput "Image directory not found: $ImagePath" -Type Warning
        return
    }
    
    Write-ColorOutput "Validating lot images in: $ImagePath" -Type Processing
    
    $found = 0
    $missing = @()
    
    foreach ($row in $Data) {
        if ($row.Lot) {
            $lotNumber = [string]$row.Lot
            $imagePath = Join-Path $ImagePath "$lotNumber.jpg"
            
            # Also check for alternative formats
            $altPaths = @(
                (Join-Path $ImagePath "$lotNumber.jpeg"),
                (Join-Path $ImagePath "$lotNumber.png"),
                (Join-Path $ImagePath "lot$lotNumber.jpg")
            )
            
            $imageFound = Test-Path $imagePath
            if (!$imageFound) {
                foreach ($alt in $altPaths) {
                    if (Test-Path $alt) {
                        $imageFound = $true
                        Write-ColorOutput "Found alternative image: $(Split-Path $alt -Leaf)" -Type Debug
                        break
                    }
                }
            }
            
            if ($imageFound) {
                $found++
            }
            else {
                $missing += $lotNumber
            }
        }
    }
    
    $script:Statistics.MissingImages = $missing
    
    # Display results
    $total = $Data.Count
    $percent = if ($total -gt 0) { [math]::Round(($found / $total) * 100, 2) } else { 0 }
    
    Write-Host "`n📷 Image Validation Results:" -ForegroundColor Cyan
    Write-Host "────────────────────────────" -ForegroundColor Gray
    Write-Host "  Total Lots:    $total" -ForegroundColor White
    Write-Host "  Images Found:  $found" -ForegroundColor Green
    Write-Host "  Images Missing: $($missing.Count)" -ForegroundColor $(if ($missing.Count -gt 0) {"Yellow"} else {"Green"})
    Write-Host "  Success Rate:  $percent%" -ForegroundColor $(if ($percent -ge 80) {"Green"} elseif ($percent -ge 50) {"Yellow"} else {"Red"})
    
    if ($missing.Count -gt 0 -and $missing.Count -le 10) {
        Write-Host "`n  Missing images for lots:" -ForegroundColor Yellow
        $missing | ForEach-Object { Write-Host "    - $_" -ForegroundColor Gray }
    }
    elseif ($missing.Count -gt 10) {
        Write-Host "`n  Missing images for $($missing.Count) lots (showing first 10):" -ForegroundColor Yellow
        $missing | Select-Object -First 10 | ForEach-Object { Write-Host "    - $_" -ForegroundColor Gray }
    }
}
#endregion

#region Export Functions
function Export-ToCSV {
    <#
    .SYNOPSIS
        Exports data to CSV with enhanced formatting
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced export handling
    #>
    param(
        $Data,
        [string]$OutputPath
    )
    
    try {
        if (!$Data -or $Data.Count -eq 0) {
            throw "No data to export"
        }
        
        Write-ColorOutput "Exporting to CSV..." -Type Processing
        
        # Clean sheet name for filename
        $cleanSheetName = if ($SheetName) {
            $SheetName -replace '[^\w\-]', '_'
        } else {
            "Sheet"
        }
        
        # Generate filename
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $baseFileName = [System.IO.Path]::GetFileNameWithoutExtension($ExcelPath)
        $csvFileName = "${baseFileName}_${cleanSheetName}_${timestamp}.csv"
        $csvPath = Join-Path $OutputPath $csvFileName
        
        # Ensure all required columns are in the correct order
        $orderedData = @()
        foreach ($row in $Data) {
            $newRow = [PSCustomObject]@{}
            
            # Add required columns first
            foreach ($col in $script:RequiredColumns) {
                if ($row.PSObject.Properties.Name -contains $col) {
                    $newRow | Add-Member -NotePropertyName $col -NotePropertyValue $row.$col
                }
            }
            
            # Add optional columns
            foreach ($col in $script:OptionalColumns) {
                if ($row.PSObject.Properties.Name -contains $col) {
                    $value = if ($null -eq $row.$col) { "" } else { $row.$col }
                    $newRow | Add-Member -NotePropertyName $col -NotePropertyValue $value
                }
            }
            
            # Add any remaining columns
            foreach ($prop in $row.PSObject.Properties) {
                if ($prop.Name -notin $script:RequiredColumns -and $prop.Name -notin $script:OptionalColumns) {
                    $newRow | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $prop.Value -ErrorAction SilentlyContinue
                }
            }
            
            $orderedData += $newRow
        }
        
        # Export to CSV
        $orderedData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        
        # Verify export
        if (!(Test-Path $csvPath)) {
            throw "CSV file was not created"
        }
        
        $fileInfo = Get-Item $csvPath
        if ($fileInfo.Length -eq 0) {
            throw "CSV file is empty"
        }
        
        $script:Statistics.ExportedRows = $orderedData.Count
        
        Write-ColorOutput "CSV exported successfully: $csvFileName" -Type Success
        Write-Host "  Full path: $csvPath" -ForegroundColor Gray
        Write-Host "  File size: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Gray
        
        return $csvPath
    }
    catch {
        Write-ColorOutput "Failed to export CSV: $_" -Type Error
        throw
    }
}

function Show-ExportSummary {
    <#
    .SYNOPSIS
        Displays comprehensive export summary
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced summary display
    #>
    param([string]$CSVPath)
    
    $duration = (Get-Date) - $script:Statistics.StartTime
    
    Write-Host "`n" -NoNewline
    Show-Banner "EXPORT SUMMARY" "Green"
    
    Write-Host "`n📊 Processing Statistics:" -ForegroundColor Cyan
    Write-Host "────────────────────────" -ForegroundColor Gray
    Write-Host "  Source File:         $(Split-Path $ExcelPath -Leaf)" -ForegroundColor White
    Write-Host "  Sheet Name:          $SheetName" -ForegroundColor White
    Write-Host "  Total Rows Read:     $($script:Statistics.TotalRows)" -ForegroundColor White
    Write-Host "  Rows Filtered Out:   $($script:Statistics.FilteredRows)" -ForegroundColor $(if($script:Statistics.FilteredRows -gt 0){"Yellow"}else{"Gray"})
    Write-Host "  Rows Exported:       $($script:Statistics.ExportedRows)" -ForegroundColor Green
    Write-Host "  Processing Time:     $($duration.ToString('mm\:ss'))" -ForegroundColor White
    
    if ($script:Statistics.ColumnsAdded.Count -gt 0) {
        Write-Host "`n➕ Columns Added:" -ForegroundColor Yellow
        $script:Statistics.ColumnsAdded | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Gray
        }
    }
    
    if ($script:Statistics.Warnings.Count -gt 0) {
        Write-Host "`n⚠️  Warnings:" -ForegroundColor Yellow
        $script:Statistics.Warnings | ForEach-Object {
            Write-Host "  - $_" -ForegroundColor Gray
        }
    }
    
    if ($DiagnosticMode -and $script:Statistics.DiagnosticInfo.Count -gt 0) {
        Write-Host "`n🔍 Diagnostic Information:" -ForegroundColor Magenta
        foreach ($key in $script:Statistics.DiagnosticInfo.Keys) {
            Write-Host "  $key`: $($script:Statistics.DiagnosticInfo[$key])" -ForegroundColor Gray
        }
    }
    
    Write-Host "`n✅ Export Complete!" -ForegroundColor Green
    Write-Host "────────────────────" -ForegroundColor Gray
    Write-Host "  Output File: $(Split-Path $CSVPath -Leaf)" -ForegroundColor White
    Write-Host "  Directory:   $(Split-Path $CSVPath -Parent)" -ForegroundColor Gray
    
    Write-Host "`n💡 Next Steps:" -ForegroundColor Cyan
    Write-Host "  1. Review the exported CSV file" -ForegroundColor White
    Write-Host "  2. Ensure lot images are in the correct directory" -ForegroundColor White
    Write-Host "  3. Run the main logistics email script:" -ForegroundColor White
    Write-Host "     .\Generate-LogisticsEmail.ps1 -CSVPath `"$CSVPath`" -ImageDirectory `"<YourImagePath>`"" -ForegroundColor Gray
}

function New-SampleImageDirectory {
    <#
    .SYNOPSIS
        Creates sample image directory structure
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced directory creation
    #>
    param(
        $Data,
        [string]$BasePath
    )
    
    if (!$Data -or $Data.Count -eq 0) {
        Write-ColorOutput "No data available for creating sample structure" -Type Warning
        return
    }
    
    Write-Host "`n📁 Would you like to create a sample image directory structure?" -ForegroundColor Cyan
    $create = Read-Host "This will create placeholder .txt files for missing images (Y/N)"
    
    if ($create -eq 'Y') {
        $imagePath = Join-Path $BasePath "LotImages"
        if (!(Test-Path $imagePath)) {
            New-Item -ItemType Directory -Path $imagePath -Force | Out-Null
        }
        
        $created = 0
        foreach ($row in $Data) {
            if ($row.Lot) {
                $placeholderPath = Join-Path $imagePath "$($row.Lot).txt"
                if (!(Test-Path $placeholderPath)) {
                    $content = @"
Placeholder for Lot $($row.Lot)
Description: $($row.Description)
Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

Replace this file with $($row.Lot).jpg
"@
                    $content | Out-File -FilePath $placeholderPath
                    $created++
                }
            }
        }
        
        Write-ColorOutput "Sample structure created in: $imagePath" -Type Success
        Write-Host "  Created $created placeholder files" -ForegroundColor Gray
        Write-Host "  Replace .txt files with actual .jpg images" -ForegroundColor Yellow
    }
}
#endregion

#region Interactive Mode
function Start-InteractiveMode {
    <#
    .SYNOPSIS
        Enhanced interactive mode with better user experience
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.1.0
        Change Date: 2025-01-07
        Change Purpose: Enhanced interactive experience
    #>
    
    Clear-Host
    Show-Banner "EXCEL TO CSV CONVERTER - INTERACTIVE MODE" "Cyan"
    
    Write-Host "`nThis wizard will guide you through the export process." -ForegroundColor White
    Write-Host "────────────────────────────────────────────────────" -ForegroundColor Gray
    
    # File selection
    if (!$ExcelPath) {
        Write-Host "`n📂 Select Excel File:" -ForegroundColor Cyan
        Write-Host "  1. Browse for file (recommended)" -ForegroundColor White
        Write-Host "  2. Enter path manually" -ForegroundColor White
        
        $choice = Read-Host "Select option (1 or 2)"
        
        if ($choice -eq "1") {
            Add-Type -AssemblyName System.Windows.Forms
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
            $OpenFileDialog.Title = "Select Excel File to Convert"
            
            if ($OpenFileDialog.ShowDialog() -eq 'OK') {
                $script:ExcelPath = $OpenFileDialog.FileName
                Write-ColorOutput "Selected: $(Split-Path $ExcelPath -Leaf)" -Type Success
            }
            else {
                Write-ColorOutput "No file selected. Exiting." -Type Warning
                return
            }
        }
        else {
            $script:ExcelPath = Read-Host "Enter full path to Excel file"
            if (!(Test-Path $script:ExcelPath)) {
                Write-ColorOutput "File not found. Exiting." -Type Error
                return
            }
        }
    }
    
    # Diagnostic check option
    Write-Host "`n🔍 Run diagnostic check first?" -ForegroundColor Cyan
    $runDiag = Read-Host "This can help identify issues with the file (Y/N) [N]"
    if ($runDiag -eq 'Y') {
        $script:DiagnosticMode = $true
        Test-ExcelFile -Path $script:ExcelPath
        Write-Host "`nPress Enter to continue..." -ForegroundColor Gray
        Read-Host
    }
    
    # Output directory
    Write-Host "`n📁 Output Directory:" -ForegroundColor Cyan
    Write-Host "  Current: $OutputDirectory" -ForegroundColor White
    $changeDir = Read-Host "  Change directory? (Y/N) [N]"
    if ($changeDir -eq 'Y') {
        $newDir = Read-Host "  Enter new directory path"
        if ($newDir -and (Test-Path $newDir -IsValid)) {
            $script:OutputDirectory = $newDir
            if (!(Test-Path $script:OutputDirectory)) {
                New-Item -ItemType Directory -Path $script:OutputDirectory -Force | Out-Null
            }
            Write-ColorOutput "Output directory changed" -Type Success
        }
        else {
            Write-ColorOutput "Invalid path. Using default." -Type Warning
        }
    }
    
    # Filtering options
    Write-Host "`n🔍 Filtering Options:" -ForegroundColor Cyan
    
    $filterByAddress = Read-Host "  Filter by pickup address? (Y/N) [N]"
    if ($filterByAddress -eq 'Y') {
        $script:FilterAddress = Read-Host "  Enter address filter (partial match)"
    }
    
    $excludeZero = Read-Host "  Exclude zero quantity items? (Y/N) [N]"
    if ($excludeZero -eq 'Y') {
        $script:ExcludeZeroQuantity = $true
    }
    
    # Validation options
    Write-Host "`n✔️ Validation Options:" -ForegroundColor Cyan
    
    $addColumns = Read-Host "  Auto-add missing required columns? (Y/N) [Y]"
    if ($addColumns -ne 'N') {
        $script:AddMissingColumns = $true
    }
    
    $validateImg = Read-Host "  Validate lot images exist? (Y/N) [N]"
    if ($validateImg -eq 'Y') {
        $script:ValidateImages = Read-Host "  Enter image directory path"
        if (!(Test-Path $script:ValidateImages)) {
            Write-ColorOutput "Image directory not found. Skipping validation." -Type Warning
            $script:ValidateImages = $null
        }
    }
    
    # Export options
    Write-Host "`n📤 Export Options:" -ForegroundColor Cyan
    
    $openFile = Read-Host "  Open CSV after export? (Y/N) [Y]"
    if ($openFile -ne 'N') {
        $script:OpenAfterExport = $true
    }
    
    Write-Host "`n" -NoNewline
    Show-Banner "STARTING EXPORT PROCESS" "Green"
    Write-Host ""
}
#endregion

#region Main Execution
try {
    # Run diagnostic mode if specified
    if ($DiagnosticMode -and !$InteractiveMode) {
        Write-ColorOutput "Running in diagnostic mode..." -Type Debug
        Test-ExcelFile -Path $ExcelPath
        Write-Host ""
    }
    
    # Run interactive mode if specified
    if ($InteractiveMode) {
        Start-InteractiveMode
    }
    
    # Verify output directory
    if (!(Test-Path $OutputDirectory)) {
        Write-ColorOutput "Creating output directory: $OutputDirectory" -Type Processing
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    }
    
    # Get available sheets
    $sheets = Get-ExcelSheets -Path $ExcelPath
    Write-ColorOutput "Found $($sheets.Count) sheet(s) in Excel file" -Type Info
    
    # Select sheet
    if (!$SheetName) {
        # Force sheets to be an array to prevent string indexing
        [array]$sheets = @($sheets)
        
        if ($sheets.Count -eq 1) {
            $SheetName = $sheets[0]
            Write-ColorOutput "Using sheet: $SheetName" -Type Info
        }
        else {
            $SheetName = Select-Sheet -Sheets $sheets
        }
    }
    elseif ($SheetName -notin $sheets) {
        Write-ColorOutput "Sheet '$SheetName' not found!" -Type Error
        Write-Host "Available sheets: $($sheets -join ', ')" -ForegroundColor Yellow
        $SheetName = Select-Sheet -Sheets $sheets
    }
    
    # Import data
    $excelData = Import-ExcelData -Path $ExcelPath -Sheet $SheetName
    
    if ($excelData -and $excelData.Count -gt 0) {
        # Validate columns
        Test-RequiredColumns -Data $excelData
        
        # Apply filters
        $processedData = Apply-DataFilters -Data $excelData
        
        if ($processedData -and $processedData.Count -gt 0) {
            # Validate images if requested
            if ($ValidateImages) {
                Test-LotImages -Data $processedData -ImagePath $ValidateImages
            }
            
            # Export to CSV
            $csvPath = Export-ToCSV -Data $processedData -OutputPath $OutputDirectory
            
            # Show summary
            Show-ExportSummary -CSVPath $csvPath
            
            # Create sample directory structure if needed
            if ($script:Statistics.MissingImages.Count -gt 0) {
                New-SampleImageDirectory -Data $processedData -BasePath $OutputDirectory
            }
            
            # Open CSV if requested
            if ($OpenAfterExport) {
                Write-Host "`n📄 Opening CSV file..." -ForegroundColor Cyan
                Start-Process $csvPath
            }
            
            # Return result object
            [PSCustomObject]@{
                Success = $true
                CSVPath = $csvPath
                RowsExported = $script:Statistics.ExportedRows
                SheetName = $SheetName
                Statistics = $script:Statistics
            }
        }
        else {
            throw "No data available after processing"
        }
    }
    else {
        throw "No data could be imported from the Excel file"
    }
}
catch {
    Write-ColorOutput "Export failed: $_" -Type Error
    Write-Host "`nError Details:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Gray
    
    if ($_.ScriptStackTrace) {
        Write-Host "`nStack Trace:" -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    }
    
    Write-Host "`n💡 Troubleshooting Tips:" -ForegroundColor Yellow
    Write-Host "  1. Ensure the Excel file is not open in another program" -ForegroundColor Gray
    Write-Host "  2. Check if the sheet has data (not just formatting)" -ForegroundColor Gray
    Write-Host "  3. Try running with -DiagnosticMode flag for detailed analysis" -ForegroundColor Gray
    Write-Host "  4. Try with -ForceRawImport flag if standard import fails" -ForegroundColor Gray
    Write-Host "  5. Ensure you have the latest ImportExcel module:" -ForegroundColor Gray
    Write-Host "     Update-Module ImportExcel -Force" -ForegroundColor Cyan
    
    # Return error object
    [PSCustomObject]@{
        Success = $false
        Error = $_.Exception.Message
        SheetName = $SheetName
        DiagnosticInfo = $script:Statistics.DiagnosticInfo
    }
    
    exit 1
}
#endregion
