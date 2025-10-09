<#
.SYNOPSIS
    Automated Logistics Quote Email Generator
    
.DESCRIPTION
    This script automates the creation of formatted HTML emails for logistics companies
    requesting freight quotes. It processes auction data from CSV files, generates
    PDF attachments with item images, and provides comprehensive logging and reporting.
    
.PARAMETER CSVPath
    Path to the CSV file containing auction lot data
    
.PARAMETER ImageDirectory
    Directory containing lot images (format: lotnumber.jpg)
    
.PARAMETER OutputDirectory
    Directory where HTML emails and PDFs will be saved
    
.PARAMETER LogDirectory
    Directory for log files (default: .\Logs)
    
.PARAMETER ShowDashboard
    Switch to display the interactive dashboard after processing
    
.EXAMPLE
    .\Generate-LogisticsEmail.ps1 -CSVPath ".\auction_data.csv" -ImageDirectory ".\LotImages" -OutputDirectory ".\Output"
    
.EXAMPLE
    .\Generate-LogisticsEmail.ps1 -CSVPath ".\data.csv" -ImageDirectory ".\Images" -ShowDashboard
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 2025-01-07
    Version: 1.0.0
    Change Date: 
    Change Purpose: Initial Release

.LINK
    https://github.com/Azure-Innovators/LogisticsAutomation
    
.COMPONENT
    Requires PowerShell 5.1 or higher
    Requires PSWritePDF module for PDF generation
    Requires ImportExcel module for Excel processing
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$CSVPath,
    
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Container})]
    [string]$ImageDirectory,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputDirectory = ".\Output",
    
    [Parameter(Mandatory=$false)]
    [string]$LogDirectory = ".\Logs",
        
    [Parameter(Mandatory=$false)]
    [ValidateRange(1,20)]
    [int]$MaxImagesPerLot = 3,

    [switch]$CreateOutlookDraft,

    [switch]$ShowDashboard
    # NEW: Config/Template-based email mode
    [Parameter(Mandatory=$false)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$ConfigPath,

    [Parameter(Mandatory=$false)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$TemplatePath,

    [Parameter(Mandatory=$false)]
    [string]$TemplateTo,

    [Parameter(Mandatory=$false)]
    [string]$RequesterName = $env:USERNAME,

    [Parameter(Mandatory=$false)]
    [string]$RequesterPhone = "",

    [switch]$OpenAfterCreateTemplate

)

#region Module Requirements
# Check and install required modules
$RequiredModules = @('PSWritePDF', 'ImportExcel')
foreach ($Module in $RequiredModules) {
    if (!(Get-Module -ListAvailable -Name $Module)) {
        Write-Host "Installing module: $Module" -ForegroundColor Yellow
        Install-Module -Name $Module -Force -AllowClobber -Scope CurrentUser
    }
    Import-Module $Module -Force
}
#endregion

#region Initialize Logging
function Initialize-Logging {
    <#
    .SYNOPSIS
        Initializes the logging infrastructure
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$LogDir
    )
    
    if (!(Test-Path $LogDir)) {
        New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    }
    
    $script:LogFile = Join-Path $LogDir ("LogisticsEmail_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")
    $script:ErrorLogFile = Join-Path $LogDir ("LogisticsEmail_Errors_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")
    $script:ProcessingStats = @{
        StartTime = Get-Date
        TotalLots = 0
        ProcessedLots = 0
        FailedLots = 0
        ImagesFound = 0
        ImagesMissing = 0
        PDFsGenerated = 0
        EmailsGenerated = 0
        DataSources = @()
    }
    
    Write-Log "=== Logistics Email Automation Started ===" -Level "INFO"
    Write-Log "Script Version: 1.0.0" -Level "INFO"
    Write-Log "User: $env:USERNAME" -Level "INFO"
    Write-Log "Machine: $env:COMPUTERNAME" -Level "INFO"
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes detailed log entries
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "DEBUG", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    # Write to console with color coding
    switch ($Level) {
        "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $LogEntry -ForegroundColor Green }
        "DEBUG"   { Write-Host $LogEntry -ForegroundColor Gray }
        default   { Write-Host $LogEntry }
    }
    
    # Write to log file
    Add-Content -Path $script:LogFile -Value $LogEntry
    
    # Write errors to separate error log
    if ($Level -eq "ERROR") {
        Add-Content -Path $script:ErrorLogFile -Value $LogEntry
    }
}
#endregion


#region Template Engine (ConfigPath mode)
function Get-Json {
    param([Parameter(Mandatory)][string]$Path)
    if (!(Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -Depth 50
}

function Get-Template {
    param([Parameter(Mandatory)][string]$Path)
    if (!(Test-Path -LiteralPath $Path)) { throw "Template not found: $Path" }
    Get-Content -LiteralPath $Path -Raw
}

function Get-ValueByPath {
    param([Parameter(Mandatory)]$Root, [Parameter(Mandatory)][string]$Path)
    $node = $Root
    foreach ($part in ($Path -split '\.')) {
        if ($null -eq $node) { return $null }
        if ($part -match '^(?<name>[^\[]+)\[(?<idx>\d+)\]$') {
            $name = $Matches.name; $idx = [int]$Matches.idx
            if ($node.PSObject.Properties[$name]) { $node = $node.$name }
            elseif ($node -is [hashtable] -and $node.ContainsKey($name)) { $node = $node[$name] }
            else { return $null }
            if ($node -is [System.Collections.IList] -and $idx -lt $node.Count) { $node = $node[$idx] } else { return $null }
            continue
        }
        if ($node -is [hashtable]) {
            if ($node.ContainsKey($part)) { $node = $node[$part] } else { return $null }
        } else {
            if ($node.PSObject.Properties[$part]) { $node = $node.$part } else { return $null }
        }
    }
    return $node
}

function Format-ForEmail {
    param($Value)
    if ($null -eq $Value) { return "" }
    switch ($Value.GetType().Name) {
        {$_ -in @("Object[]","List`1","ArrayList")} {
            ($Value | ForEach-Object { $s = (Format-ForEmail $_); if ($s) { "- $s" } }) -join [Environment]::NewLine
        }
        "DateTime" { $Value.ToString("dddd, MMMM d, yyyy h:mm tt") }
        default {
            if ($Value -is [PSCustomObject] -or $Value -is [hashtable]) {
                ($Value | ConvertTo-Json -Depth 10 -Compress)
            } else { "$Value" }
        }
    }
}

function Expand-Template {
    param([string]$Template, $Model)
    $pattern = '\{\{\s*([^\}]+?)\s*\}\}'
    [regex]::Replace($Template, $pattern, { param($m)
        $path = $m.Groups[1].Value.Trim()
        $val  = Get-ValueByPath -Root $Model -Path $path
        (Format-ForEmail $val)
    })
}

function Extract-SubjectAndBody {
    param([string]$Expanded)
    $subject = $null; $body = $Expanded
    $lines = $Expanded -split '\r?\n'
    for ($i=0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()
        if ($line -match '^(?i)Subject:\s*(.+)$') {
            $subject = $Matches[1].Trim()
            $lines.RemoveAt($i)
            $body = ($lines -join [Environment]::NewLine).TrimStart()
            break
        } elseif ($line -ne "") { break }
    }
    [pscustomobject]@{ Subject=$subject; Body=$body }
}
#endregion

#region Data Processing Functions
function Import-AuctionData {
    <#
    .SYNOPSIS
        Imports and validates auction data from CSV
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$CSVPath
    )
    
    try {
        Write-Log "Importing CSV data from: $CSVPath" -Level "INFO"
        $script:ProcessingStats.DataSources += $CSVPath
        
        $Data = Import-Csv -Path $CSVPath
        $script:ProcessingStats.TotalLots = $Data.Count
        
        Write-Log "Successfully imported $($Data.Count) lots" -Level "SUCCESS"
        
        # Validate required columns
        $RequiredColumns = @('Lot', 'Description', 'Address')
        $MissingColumns = $RequiredColumns | Where-Object { $_ -notin $Data[0].PSObject.Properties.Name }
        
        if ($MissingColumns) {
            throw "Missing required columns: $($MissingColumns -join ', ')"
        }
        
        return $Data
    }
    catch {
        Write-Log "Failed to import CSV: $_" -Level "ERROR"
        throw
    }
}

function Get-LotImages {
    <#
    .SYNOPSIS
        Retrieves images for specified lot numbers with intelligent selection
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.2.1
        Change Date: 2025-01-07
        Change Purpose: Fixed numerical sorting for image selection
    #>
    param (
        [array]$Lots,
        [string]$ImageDir,
        [int]$MaxImagesPerLot = 3
    )
    
    $ImageList = @()
    
    foreach ($Lot in $Lots) {
        $LotNumber = $Lot.Lot
        $AllFoundImages = @()
        
        # Priority 1: Primary image (lotnumber.jpg)
        $PrimaryImagePath = Join-Path $ImageDir "$LotNumber.jpg"
        if (Test-Path $PrimaryImagePath) {
            $AllFoundImages += @{
                Path = $PrimaryImagePath
                Priority = 1
                SortOrder = 0
                Type = "Primary"
            }
        }
        
        # Priority 2: Numbered images (lotnumber-2.jpg, lotnumber-3.jpg, etc.)
        # Find all numbered images first
        $NumberedImages = @()
        for ($i = 2; $i -le 30; $i++) {
            $NumberedImagePath = Join-Path $ImageDir "$LotNumber-$i.jpg"
            if (Test-Path $NumberedImagePath) {
                $NumberedImages += @{
                    Path = $NumberedImagePath
                    Priority = 2
                    SortOrder = $i  # Use the actual number for sorting
                    Type = "Image$i"
                }
            }
        }
        
        # Add numbered images sorted by their actual number
        $AllFoundImages += $NumberedImages | Sort-Object SortOrder
        
        # Priority 3: Lettered variants (for different angles/views)
        foreach ($letter in @('a','b','c','d','e','f')) {
            $LetterImagePath = Join-Path $ImageDir "$LotNumber-$letter.jpg"
            if (Test-Path $LetterImagePath) {
                $AllFoundImages += @{
                    Path = $LetterImagePath
                    Priority = 3
                    SortOrder = 100 + [int][char]$letter  # Convert letter to number for sorting
                    Type = "Variant-$letter"
                }
            }
        }
        
        if ($AllFoundImages.Count -gt 0) {
            # Sort by priority first, then by sort order (which is now numerical)
            $SelectedImages = $AllFoundImages | 
                Sort-Object Priority, SortOrder | 
                Select-Object -First $MaxImagesPerLot
            
            $ImagePaths = $SelectedImages | ForEach-Object { $_.Path }
            
            Write-Log "Lot $LotNumber`: Found $($AllFoundImages.Count) images, including $($ImagePaths.Count) (max: $MaxImagesPerLot)" -Level "DEBUG"
            
            # Log which images were selected for debugging
            $selectedNames = $SelectedImages | ForEach-Object { Split-Path $_.Path -Leaf }
            Write-Log "  Selected: $($selectedNames -join ', ')" -Level "DEBUG"

            if ($AllFoundImages.Count -gt $MaxImagesPerLot) {
                Write-Log "Lot $LotNumber has $($AllFoundImages.Count) images, selected: $(($SelectedImages.Type) -join ', ')" -Level "INFO"
            }
            
            $script:ProcessingStats.ImagesFound += $ImagePaths.Count
            
            $ImageList += @{
                LotNumber = $LotNumber
                Description = $Lot.Description
                ImagePaths = $ImagePaths
                ImageCount = $ImagePaths.Count
                TotalFound = $AllFoundImages.Count
                FileSize = ($ImagePaths | ForEach-Object { (Get-Item $_).Length } | Measure-Object -Sum).Sum
            }
            
            if ($AllFoundImages.Count -gt $MaxImagesPerLot) {
                Write-Log "Lot $LotNumber has $($AllFoundImages.Count) images, selected: $(($SelectedImages.Type) -join ', ')" -Level "INFO"
            }
        }
        else {
            Write-Log "No images found for Lot $LotNumber" -Level "WARNING"
            $script:ProcessingStats.ImagesMissing++
        }
    }
    
    $totalIncluded = ($ImageList.ImageCount | Measure-Object -Sum).Sum
    $totalFound = ($ImageList.TotalFound | Measure-Object -Sum).Sum
    Write-Log "Including $totalIncluded of $totalFound total images found" -Level "INFO"
    
    return $ImageList
}
#endregion

#region PDF Generation
function New-LotPDF {
    <#
    .SYNOPSIS
        Creates HTML report with multiple lot images per lot
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.2.1
        Change Date: 2025-01-07
        Change Purpose: Fixed ImagePaths array handling
    #>
    param (
        [array]$Images,
        [string]$OutputPath,
        [string]$AuctionName
    )
    
    try {
        # Calculate total image count correctly
        $totalImageCount = 0
        foreach ($lot in $Images) {
            if ($lot.ImagePaths) {
                $totalImageCount += $lot.ImagePaths.Count
            }
        }
        
        Write-Log "Generating image report for $($Images.Count) lots with $totalImageCount total images" -Level "INFO"
        
        $ReportPath = Join-Path $OutputPath ("AuctionLots_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".html")
        
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Auction Lot Images - $AuctionName</title>
    <style>
        @media print {
            .page { page-break-after: always; }
            .no-print { display: none; }
        }
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            margin: 20px;
            background: #f5f5f5;
        }
        .header {
            background: #2c3e50;
            color: white;
            padding: 20px;
            text-align: center;
            margin-bottom: 20px;
        }
        .page {
            background: white;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .lot-number { 
            font-size: 24px; 
            font-weight: bold; 
            color: #2c3e50;
            margin-bottom: 10px;
        }
        .description { 
            font-size: 14px; 
            color: #666;
            margin-bottom: 15px;
        }
        .image-container {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-bottom: 10px;
        }
        .image-wrapper {
            flex: 1 1 48%;
            min-width: 300px;
        }
        img { 
            width: 100%;
            height: auto;
            border: 1px solid #ddd;
            border-radius: 5px;
        }
        .image-label {
            text-align: center;
            font-size: 12px;
            color: #666;
            margin-top: 5px;
        }
        .print-btn {
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 10px 20px;
            background: #3498db;
            color: white;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            z-index: 1000;
        }
    </style>
</head>
<body>
    <button class="print-btn no-print" onclick="window.print()">Print to PDF</button>
    
    <div class="header">
        <h1>Auction Lot Images</h1>
        <div>
            Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')<br>
            Total Lots: $($Images.Count)<br>
            Total Images: $totalImageCount
        </div>
    </div>
"@
        
        foreach ($LotInfo in $Images) {
            $html += @"
    <div class="page">
        <div class="lot-number">Lot #$($LotInfo.LotNumber)</div>
        <div class="description">$($LotInfo.Description)</div>
        <div class="image-container">
"@
            
            # Check if ImagePaths exists and is an array
            if ($LotInfo.ImagePaths) {
                $imageNum = 1
                $totalForLot = $LotInfo.ImagePaths.Count
                
                foreach ($imagePath in $LotInfo.ImagePaths) {
                    $imageLabel = if ($totalForLot -gt 1) { 
                        "Image $imageNum of $totalForLot" 
                    } else { 
                        "Lot Image" 
                    }
                    
                    # Convert to base64 for embedding
                    if ($imagePath -and (Test-Path $imagePath)) {
                        try {
                            $imageBytes = [System.IO.File]::ReadAllBytes($imagePath)
                            $imageBase64 = [System.Convert]::ToBase64String($imageBytes)
                            $imageSrc = "data:image/jpeg;base64,$imageBase64"
                        }
                        catch {
                            Write-Log "Could not embed image: $imagePath" -Level "WARNING"
                            $imageSrc = "file:///$($imagePath -replace '\\','/')"
                        }
                        
                        $html += @"
            <div class="image-wrapper">
                <img src="$imageSrc" alt="Lot $($LotInfo.LotNumber) - $imageLabel" />
                <div class="image-label">$imageLabel</div>
            </div>
"@
                        $imageNum++
                    }
                }
            }
            else {
                $html += @"
            <div class="image-wrapper">
                <div style="padding: 20px; background: #f0f0f0; text-align: center;">
                    No images available for this lot
                </div>
            </div>
"@
            }
            
            $html += @"
        </div>
    </div>
"@
        }
        
        $html += @"
</body>
</html>
"@
        
        $html | Out-File -FilePath $ReportPath -Encoding UTF8
        
        Write-Log "Image report generated successfully: $ReportPath" -Level "SUCCESS"
        Write-Log "Open the HTML file and use Print > Save as PDF for final PDF" -Level "INFO"
        $script:ProcessingStats.PDFsGenerated++
        
        # Auto-open the report
        Start-Process $ReportPath
        
        return $ReportPath
    }
    catch {
        Write-Log "Failed to generate image report: $_" -Level "ERROR"
        throw
    }
}

function ConvertTo-PDF {
    <#
    .SYNOPSIS
        Converts HTML file to PDF using Microsoft Word
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose: Generate actual PDF files
    #>
    param (
        [string]$HTMLPath,
        [string]$OutputDirectory
    )
    
    try {
        Write-Log "Converting HTML to PDF using Word" -Level "INFO"
        
        # Create Word COM object
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $false
        
        # Open HTML file
        $Document = $Word.Documents.Open($HTMLPath)
        
        # Generate PDF filename
        $PDFPath = Join-Path $OutputDirectory ("AuctionLots_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".pdf")
        
        # Save as PDF (17 = PDF format)
        $Document.SaveAs2($PDFPath, 17)
        
        # Close and cleanup
        $Document.Close()
        $Word.Quit()
        
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Document) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Log "PDF created successfully: $PDFPath" -Level "SUCCESS"
        return $PDFPath
    }
    catch {
        Write-Log "Failed to convert to PDF: $_" -Level "ERROR"
        Write-Log "Falling back to HTML report" -Level "WARNING"
        return $null
    }
}
#endregion

#region HTML Email Generation
function New-LogisticsEmailHTML {
    <#
    .SYNOPSIS
        Generates formatted HTML email matching Word template
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.1
        Change Date: 2025-01-07
        Change Purpose: Added lot numbers to item list
    #>
    param (
        [array]$LotData,
        [string]$PDFPath
    )
    
    try {
        Write-Log "Generating HTML email for $($LotData.Count) lots" -Level "INFO"
        
        # Get unique pickup address (assuming all lots from same auction)
        $PickupAddress = $LotData[0].Address.Trim()
        
        # Build items list with lot numbers and quantities
        $ItemsList = ""
        foreach ($Lot in $LotData) {
            $qtyText = if ($Lot.Quantity) { " (Qty: $($Lot.Quantity))" } else { "" }
            $ItemsList += "        <li><strong>Lot #$($Lot.Lot):</strong> $($Lot.Description)$qtyText</li>`n"
        }
        
        $HTML = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Logistics Quote Request</title>
    <style>
        body {
            font-family: Calibri, Arial, sans-serif;
            font-size: 11pt;
            line-height: 1.5;
            color: #000000;
            max-width: 700px;
            margin: 20px;
        }
        p { margin: 10px 0; }
        ul { margin-left: 30px; }
        li { margin: 5px 0; }
    </style>
</head>
<body>
    <p>Hello.</p>
    
    <p>We have an online auction that's closed with $($LotData.Count) lots we won to pick up. Here is the information:</p>
    
    <p><strong>PICKUP ADDRESS</strong>: $PickupAddress</p>
    
    <p><strong>DELIVERY ADDRESS:</strong> 1218 Lake Avenue, Ashtabula, OH 44004</p>
    
    <p><strong>AUCTION LOGISTICS CONTACT:</strong></p>
    
    <p><strong>PICKUP DATE/TIME:</strong></p>
    
    <p><strong>DELIVERY DATE/TIME:</strong></p>
    
    <p><strong>SPECIAL NOTES:</strong></p>
    
    <p>The items (pictures in the attached PDF referenced by Lot Number) are:</p>
    
    <ul>
$ItemsList
    </ul>
    
    <p>For this shipment, we believe two trucks are necessary. Please select the type of trucks based on the items being picked up, truck availability, and lowest cost.</p>
    
    <p>Please send me a quote at your earliest opportunity.</p>
    
    <p>Thank you.</p>
    
    <p>John O'Neill Sr.<br>
    AWS Solutions LLC dba JT Custom Trailers<br>
    (440) 813-6695</p>
</body>
</html>
"@
        
        Write-Log "HTML email generated successfully" -Level "SUCCESS"
        $script:ProcessingStats.EmailsGenerated++
        
        return $HTML
    }
    catch {
        Write-Log "Failed to generate HTML email: $_" -Level "ERROR"
        throw
    }
}

function New-OutlookDraftEmail {
    <#
    .SYNOPSIS
        Creates a draft email in Outlook with HTML formatting
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$HTMLContent,
        [string]$Subject = "Freight Quote Request - $(Get-Date -Format 'yyyy-MM-dd')",
        [string[]]$Attachments = @(),
        [string]$To = "",
        [switch]$Display
    )
    
    try {
        Write-Log "Creating Outlook draft email" -Level "INFO"
        
        # Create Outlook COM object
        $Outlook = New-Object -ComObject Outlook.Application
        $Mail = $Outlook.CreateItem(0)  # 0 = Mail item
        
        # Set email properties
        $Mail.Subject = $Subject
        $Mail.HTMLBody = $HTMLContent
        
        if ($To) {
            $Mail.To = $To
        }
        
        # Add attachments
        foreach ($Attachment in $Attachments) {
            if (Test-Path $Attachment) {
                $Mail.Attachments.Add($Attachment) | Out-Null
                Write-Log "Added attachment: $(Split-Path $Attachment -Leaf)" -Level "DEBUG"
            }
        }
        
        # Save as draft
        $Mail.Save()
        
        # Display if requested
        if ($Display) {
            $Mail.Display()
        }
        
        Write-Log "Outlook draft email created successfully" -Level "SUCCESS"
        
        # Release COM objects
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mail) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        return $true
    }
    catch {
        Write-Log "Failed to create Outlook email: $_" -Level "ERROR"
        Write-Log "Make sure Outlook is installed and configured" -Level "WARNING"
        return $false
    }
}
#endregion

#region Dashboard and Reporting
function Show-Dashboard {
    param (
        [hashtable]$Stats
    )
    
    # Add pause to review any errors
    Write-Host "`nPress Enter to view dashboard..." -ForegroundColor Yellow
    Read-Host
    
    $EndTime = Get-Date
    $Duration = $EndTime - $Stats.StartTime
    
    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║         LOGISTICS EMAIL AUTOMATION DASHBOARD              ║" -ForegroundColor Cyan
    Write-Host "╠════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
    Write-Host "║                                                            ║" -ForegroundColor Cyan
    
    # Processing Summary
    Write-Host "║  📊 PROCESSING SUMMARY                                    ║" -ForegroundColor Yellow
    Write-Host "║  ─────────────────────────────────────────────────────    ║" -ForegroundColor Gray
    Write-Host "║  Total Lots Processed: $($Stats.TotalLots.ToString().PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║  Images Found:         $($Stats.ImagesFound.ToString().PadLeft(10))                          ║" -ForegroundColor Green
    Write-Host "║  Images Missing:       $($Stats.ImagesMissing.ToString().PadLeft(10))                          ║" -ForegroundColor $(if ($Stats.ImagesMissing -gt 0) { "Yellow" } else { "White" })
    Write-Host "║  PDFs Generated:       $($Stats.PDFsGenerated.ToString().PadLeft(10))                          ║" -ForegroundColor Green
    Write-Host "║  Emails Generated:     $($Stats.EmailsGenerated.ToString().PadLeft(10))                          ║" -ForegroundColor Green
    Write-Host "║                                                            ║" -ForegroundColor Cyan
    
    # Performance Metrics
    Write-Host "║  ⏱️  PERFORMANCE METRICS                                   ║" -ForegroundColor Yellow
    Write-Host "║  ─────────────────────────────────────────────────────    ║" -ForegroundColor Gray
    Write-Host "║  Processing Time:      $($Duration.ToString('mm\:ss').PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║  Start Time:           $(($Stats.StartTime.ToString('HH:mm:ss')).PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║  End Time:             $(($EndTime.ToString('HH:mm:ss')).PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║                                                            ║" -ForegroundColor Cyan
    
    # Success Rate Graph - FIXED CALCULATION
    # Calculate based on lots with images vs total lots
    $LotsWithImages = if ($Stats.TotalLots -gt 0 -and $Stats.ImagesMissing -ge 0) {
        $Stats.TotalLots - $Stats.ImagesMissing
    } else { $Stats.TotalLots }
    
    $SuccessRate = if ($Stats.TotalLots -gt 0) { 
        [math]::Round(($LotsWithImages / $Stats.TotalLots) * 100, 2) 
    } else { 0 }
    
    # Ensure success rate is capped at 100%
    $SuccessRate = [math]::Min($SuccessRate, 100)
    
    Write-Host "║  📈 LOT IMAGE COVERAGE                                     ║" -ForegroundColor Yellow
    Write-Host "║  ─────────────────────────────────────────────────────    ║" -ForegroundColor Gray
    
    $BarLength = 40
    $FilledLength = [math]::Floor($SuccessRate / 100 * $BarLength)
    $FilledLength = [math]::Max(0, [math]::Min($FilledLength, $BarLength))  # Ensure between 0 and BarLength
    $EmptyLength = $BarLength - $FilledLength
    $Bar = "█" * $FilledLength + "░" * $EmptyLength
    
    Write-Host "║  [$Bar] $($SuccessRate.ToString('F2'))%  ║" -ForegroundColor $(if ($SuccessRate -ge 80) { "Green" } elseif ($SuccessRate -ge 50) { "Yellow" } else { "Red" })
    Write-Host "║  Lots with images: $LotsWithImages of $($Stats.TotalLots)                             ║" -ForegroundColor White
    
    if ($Stats.ImagesFound -gt $Stats.TotalLots) {
        $avgImages = [math]::Round($Stats.ImagesFound / $LotsWithImages, 1)
        Write-Host "║  Average images per lot: $($avgImages.ToString('F1'))                             ║" -ForegroundColor White
    }
    
    Write-Host "║                                                            ║" -ForegroundColor Cyan
    
    # Data Sources
    Write-Host "║  📁 DATA SOURCES                                           ║" -ForegroundColor Yellow
    Write-Host "║  ─────────────────────────────────────────────────────    ║" -ForegroundColor Gray
    foreach ($Source in $Stats.DataSources) {
        $SourceName = Split-Path $Source -Leaf
        if ($SourceName.Length -gt 50) {
            $SourceName = $SourceName.Substring(0, 47) + "..."
        }
        Write-Host "║  • $($SourceName.PadRight(54))  ║" -ForegroundColor White
    }
    
    Write-Host "║                                                            ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
    
    # Action Items
    if ($Stats.ImagesMissing -gt 0) {
        Write-Host "⚠️  WARNING: $($Stats.ImagesMissing) lots are missing images!" -ForegroundColor Yellow
        Write-Host "   Check the image directory for missing .jpg files" -ForegroundColor Gray
        Write-Host ""
    }
    
    Write-Host "✅ Processing Complete!" -ForegroundColor Green
    Write-Host "   Logs saved to: $script:LogFile" -ForegroundColor Gray
    Write-Host ""
}

function Export-ProcessingReport {
    <#
    .SYNOPSIS
        Exports detailed processing report
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$OutputPath,
        [hashtable]$Stats,
        [array]$LotData
    )
    
    $ReportPath = Join-Path $OutputPath ("ProcessingReport_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".html")
    
    $ReportHTML = @"
<!DOCTYPE html>
<html>
<head>
    <title>Logistics Processing Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        h1 { color: #2c3e50; }
        table { border-collapse: collapse; width: 100%; margin: 20px 0; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
        th { background-color: #3498db; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .summary { background-color: #ecf0f1; padding: 20px; border-radius: 5px; margin: 20px 0; }
        .warning { background-color: #fff3cd; padding: 10px; border-left: 4px solid #ffc107; margin: 10px 0; }
        .success { background-color: #d4edda; padding: 10px; border-left: 4px solid #28a745; margin: 10px 0; }
    </style>
</head>
<body>
    <h1>Logistics Email Automation - Processing Report</h1>
    <div class="summary">
        <h2>Summary Statistics</h2>
        <p><strong>Report Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p><strong>Total Lots Processed:</strong> $($Stats.TotalLots)</p>
        <p><strong>Images Found:</strong> $($Stats.ImagesFound)</p>
        <p><strong>Images Missing:</strong> $($Stats.ImagesMissing)</p>
        <p><strong>PDFs Generated:</strong> $($Stats.PDFsGenerated)</p>
        <p><strong>Emails Generated:</strong> $($Stats.EmailsGenerated)</p>
    </div>
    
    <h2>Lot Details</h2>
    <table>
        <tr>
            <th>Lot Number</th>
            <th>Description</th>
            <th>Address</th>
            <th>Quantity</th>
            <th>Image Status</th>
        </tr>
"@
    
    foreach ($Lot in $LotData) {
        $ImageStatus = if (Test-Path (Join-Path $ImageDirectory "$($Lot.Lot).jpg")) { 
            "✅ Found" 
        } else { 
            "❌ Missing" 
        }
        
        $ReportHTML += @"
        <tr>
            <td>$($Lot.Lot)</td>
            <td>$($Lot.Description)</td>
            <td>$($Lot.Address)</td>
            <td>$($Lot.Quantity)</td>
            <td>$ImageStatus</td>
        </tr>
"@
    }
    
    $ReportHTML += @"
    </table>
</body>
</html>
"@
    
    $ReportHTML | Out-File -FilePath $ReportPath -Encoding UTF8
    Write-Log "Processing report exported to: $ReportPath" -Level "SUCCESS"
    
    return $ReportPath
}
#endregion

#region Main Execution
try {
    # Initialize
    Initialize-Logging -LogDir $LogDirectory
    
    # Create output directory if it doesn't exist
    if (!(Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        Write-Log "Created output directory: $OutputDirectory" -Level "INFO"
    }
    
    
    # --- Config/Template mode (short-circuit the CSV/PDF pipeline) ---
    if ($PSBoundParameters.ContainsKey('ConfigPath') -and $PSBoundParameters.ContainsKey('TemplatePath')) {
        Write-Log "Config/Template mode detected. Using -ConfigPath and -TemplatePath to generate email draft." -Level "INFO"

        $config = Get-Json -Path $ConfigPath
        $meta   = @{
            requester_name  = $RequesterName
            requester_phone = $RequesterPhone
        }

        $model = [ordered]@{}
        $config.PSObject.Properties | ForEach-Object { $model[$_.Name] = $_.Value }
        $model['_meta'] = $meta

        $templateRaw = Get-Template -Path $TemplatePath
        $expanded    = Expand-Template -Template $templateRaw -Model $model
        $sb          = Extract-SubjectAndBody -Expanded $expanded

        $subject = if ($sb.Subject) { $sb.Subject }
                   elseif ($config.email_subject) { [string]$config.email_subject }
                   else { "Logistics Request" }

        $body = $sb.Body

        $to = if ($TemplateTo) { $TemplateTo }
              elseif ($config.auction_info -and $config.auction_info.logistics_contact -and $config.auction_info.logistics_contact.email) {
                  [string]$config.auction_info.logistics_contact.email
              } else { "" }

        # Convert plain text to HTML line breaks if needed
        $htmlBody =
            if ($body -match '<(html|body|p|br|div)[\s>]' -or $body -match '</') { $body }
            else { ($body -split '\r?\n' | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) } ) -join "<br/>" }

        if ($CreateOutlookDraft) {
            try {
                $null = New-OutlookDraftEmail -HTMLContent $htmlBody -Subject $subject -To $to -Display:$OpenAfterCreateTemplate
                Write-Log "Draft created via Config/Template mode (Subject: $subject)" -Level "SUCCESS"
            } catch {
                Write-Log "Failed to create draft in Config/Template mode: $($_.Exception.Message)" -Level "ERROR"
                throw
            }
        } else {
            Write-Host "`n=== SUBJECT ===`n$subject`n`n=== BODY (HTML) ===`n$htmlBody"
        }

        return
    }
# Import auction data
    Write-Log "Starting data import process" -Level "INFO"
    $AuctionData = Import-AuctionData -CSVPath $CSVPath
    
    # Process lot images with max image limit
    Write-Log "Processing lot images (max $MaxImagesPerLot per lot)" -Level "INFO"
    $LotImages = Get-LotImages -Lots $AuctionData -ImageDir $ImageDirectory -MaxImagesPerLot $MaxImagesPerLot

    # Generate image report
    $PDFFilePath = $null
    $ImageReportPath = $null

    if ($LotImages.Count -gt 0) {
        Write-Log "Generating image report with $($LotImages.Count) images" -Level "INFO"
        $ImageReportPath = New-LotPDF -Images $LotImages -OutputPath $OutputDirectory -AuctionName "Auction"
        
        # Try to convert HTML to actual PDF
        $PDFFilePath = ConvertTo-PDF -HTMLPath $ImageReportPath -OutputDirectory $OutputDirectory
        
        if (!$PDFFilePath) {
            # If PDF conversion failed, use HTML as fallback
            Write-Log "Using HTML report as attachment" -Level "INFO"
            $PDFFilePath = $ImageReportPath
        }
    }
    else {
        Write-Log "No images found, skipping PDF generation" -Level "WARNING"
    }
    
    # Generate HTML email
    Write-Log "Generating HTML email" -Level "INFO"
    $EmailHTML = New-LogisticsEmailHTML -LotData $AuctionData -PDFPath $PDFFilePath
    
    # Save HTML to file
    $HTMLFilePath = Join-Path $OutputDirectory ("LogisticsEmail_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".html")
    $EmailHTML | Out-File -FilePath $HTMLFilePath -Encoding UTF8
    Write-Log "HTML email saved to: $HTMLFilePath" -Level "SUCCESS"

    # Prepare attachments
    $AttachmentList = @()
    if ($PDFFilePath) {
        $AttachmentList += $PDFFilePath
    }

    # Check if Outlook draft should be created
    if ($CreateOutlookDraft) {
        $OutlookCreated = New-OutlookDraftEmail -HTMLContent $EmailHTML -Attachments $AttachmentList -Display
    }
    else {
        $OutlookCreated = $false
        Write-Log "Outlook draft creation skipped (use -CreateOutlookDraft to enable)" -Level "INFO"
    }

    # Fallback to browser if Outlook wasn't created
    if (-not $OutlookCreated) {
        Write-Log "Opening HTML in default browser" -Level "INFO"
        Start-Process $HTMLFilePath
        
        if ($CreateOutlookDraft) {
            # Only show manual instructions if Outlook was attempted but failed
            Write-Host "`n📧 Manual Email Creation Required:" -ForegroundColor Yellow
            Write-Host "  1. Open the HTML file in your browser" -ForegroundColor White
            Write-Host "  2. Press Ctrl+A to select all" -ForegroundColor White
            Write-Host "  3. Press Ctrl+C to copy" -ForegroundColor White
            Write-Host "  4. In Outlook, create new email" -ForegroundColor White
            Write-Host "  5. In the email body, press Ctrl+V to paste" -ForegroundColor White
            Write-Host "  6. Attach the PDF file manually" -ForegroundColor White
        }
    }

    # Generate processing report
    $ReportPath = Export-ProcessingReport -OutputPath $OutputDirectory -Stats $script:ProcessingStats -LotData $AuctionData
    
    # Update final statistics
    $script:ProcessingStats.ProcessedLots = $AuctionData.Count
    
    # Show dashboard if requested
    if ($ShowDashboard) {
        Show-Dashboard -Stats $script:ProcessingStats
    }
    
    Write-Log "=== Logistics Email Automation Completed Successfully ===" -Level "SUCCESS"
    Write-Log "Output files saved to: $OutputDirectory" -Level "INFO"
    
    # Return summary object
    [PSCustomObject]@{
        Success = $true
        HTMLFile = $HTMLFilePath
        PDFFile = $PDFFilePath
        ReportFile = $ReportPath
        Statistics = $script:ProcessingStats
        LogFile = $script:LogFile
    }
}
catch {
    Write-Log "Fatal error: $_" -Level "ERROR"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    
    # Return error object
    [PSCustomObject]@{
        Success = $false
        Error = $_.Exception.Message
        LogFile = $script:LogFile
        ErrorLogFile = $script:ErrorLogFile
    }
    
    throw
}
finally {
    Write-Log "Script execution completed" -Level "INFO"
    Write-Log "Total execution time: $((Get-Date) - $script:ProcessingStats.StartTime)" -Level "INFO"
}
#endregion
