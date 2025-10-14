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
    Directory where HTML emails and PDFs will be saved (defaults to config file location\Output)
    
.PARAMETER LogDirectory
    Directory for log files (defaults to config file location\Logs)
    
.PARAMETER ConfigPath
    Path to JSON configuration file with auction details
    
.PARAMETER ShowDashboard
    Switch to display the interactive dashboard after processing
    
.EXAMPLE
    .\Generate-LogisticsEmail.ps1 -CSVPath ".\auction_data.csv" -ImageDirectory ".\LotImages" -OutputDirectory ".\Output"
    
.EXAMPLE
    .\Generate-LogisticsEmail.ps1 -CSVPath ".\data.csv" -ImageDirectory ".\Images" -ConfigPath ".\config.json" -ShowDashboard
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 01/07/2025
    Version: 1.5.4
    Change Date: 10/10/2025
    Change Purpose: Fixed apostrophe in username (O'Neill) breaking PDF conversion

.CHANGELOG
    1.5.4 - 10/10/2025 - Helper uses C:\Temp instead of user temp folder
                       - Fixes apostrophe in username breaking command-line parsing
                       - Solves Edge/Chrome "exit code 1" with special characters in paths
    1.5.3 - 10/10/2025 - Helper script now handles OneDrive OUTPUT paths too
                       - Creates PDF in temp, then copies to OneDrive
    1.5.2 - 10/10/2025 - Helper script now handles OneDrive paths automatically
                       - Creates temp copies for conversion (fixes Edge/Chrome issues)
    1.5.1 - 10/10/2025 - Added detailed PDF conversion output for debugging
                       - Removed -Quiet flag to show all conversion steps
                       - Opens Explorer to PDF location on success
    1.5.0 - 10/10/2025 - Integrated Convert-HTMLtoPDF.ps1 helper script
                       - Automatic PDF conversion using Foxit, Edge, or Chrome
                       - Removed unreliable Word COM automation completely
                       - Fast, reliable PDF creation in 3-5 seconds
    1.4.0 - 10/10/2025 - Word PDF conversion now DISABLED by default (prevents hanging)
                       - Added -TryWordPDFConversion parameter for optional Word conversion
                       - Provides clear manual PDF creation instructions instead
    1.3.1 - 10/09/2025 - Added timeout protection to Word PDF conversion (45s default)
                       - Improved error handling for Word COM automation
                       - Better process cleanup to prevent hanging
    1.3.0 - 10/09/2025 - Logs now stored in Logs subdirectory of config file location
                       - Output now stored in Output subdirectory of config file location
    1.2.0 - 10/09/2025 - Fixed special notes bullet encoding (now uses HTML list)
                       - Added prominent display of HTML file absolute path
    1.1.1 - 10/09/2025 - Fixed Word PDF conversion path issues (now uses absolute paths)
                       - Fixed Outlook attachment path issues (now uses absolute paths)
    1.1.0 - 10/09/2025 - Modified New-LogisticsEmailHTML to use config file data
    1.0.0 - 01/07/2025 - Initial Release

.LINK
    https://github.com/JONeillSr/Shipping-Management
    
.COMPONENT
    Requires PowerShell 5.1 or higher
    Requires PSWritePDF module for PDF generation
    Requires ImportExcel module for Excel processing
    Requires Convert-HTMLtoPDF.ps1 helper script (optional but recommended)
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

    [Parameter(Mandatory=$false)]
    [switch]$CreateOutlookDraft,

    [Parameter(Mandatory=$false)]
    [switch]$ShowDashboard,
    
    [Parameter(Mandatory=$false)]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$ConfigPath,

    [Parameter(Mandatory=$false)]
    [string]$RequesterName = $env:USERNAME,

    [Parameter(Mandatory=$false)]
    [string]$RequesterPhone = ""
)

#region Load Helper Scripts
# Load PDF conversion helper (optional but highly recommended)
$helperScript = Join-Path $PSScriptRoot "Convert-HTMLtoPDF.ps1"
$script:PDFHelperAvailable = $false

if (Test-Path $helperScript) {
    try {
        . $helperScript
        $script:PDFHelperAvailable = $true
        Write-Host "✓ Loaded PDF conversion helper (Foxit/Edge/Chrome support)" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️  Failed to load PDF helper: $_" -ForegroundColor Yellow
        Write-Host "   PDF conversion will not be available" -ForegroundColor Gray
    }
}
else {
    Write-Host "⚠️  PDF helper not found: $helperScript" -ForegroundColor Yellow
    Write-Host "   Download from: https://github.com/JONeillSr/Shipping-Management" -ForegroundColor Gray
    Write-Host "   Continuing without automatic PDF conversion..." -ForegroundColor Gray
}
#endregion

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
        Create Date: 01/07/2025
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
    Write-Log "Script Version: 1.5.4" -Level "INFO"
    Write-Log "User: $env:USERNAME" -Level "INFO"
    Write-Log "Machine: $env:COMPUTERNAME" -Level "INFO"
    Write-Log "PDF Helper Available: $script:PDFHelperAvailable" -Level "INFO"
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes detailed log entries
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 01/07/2025
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

#region Configuration Loading
function Get-AuctionConfig {
    <#
    .SYNOPSIS
        Loads and parses JSON configuration file
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 10/09/2025
        Version: 1.0.0
        Change Date: 
        Change Purpose: Load config for email generation
    #>
    param(
        [string]$ConfigPath
    )
    
    if (-not $ConfigPath -or -not (Test-Path $ConfigPath)) {
        Write-Log "No config file provided or file not found" -Level "DEBUG"
        return $null
    }
    
    try {
        Write-Log "Loading configuration from: $ConfigPath" -Level "INFO"
        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        Write-Log "Configuration loaded successfully" -Level "SUCCESS"
        return $config
    }
    catch {
        Write-Log "Failed to load configuration: $_" -Level "ERROR"
        return $null
    }
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
        Create Date: 01/07/2025
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
        $RequiredColumns = @('Lot', 'Description')
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
        Create Date: 01/07/2025
        Version: 1.2.1
        Change Date: 01/07/2025
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
        $NumberedImages = @()
        for ($i = 2; $i -le 30; $i++) {
            $NumberedImagePath = Join-Path $ImageDir "$LotNumber-$i.jpg"
            if (Test-Path $NumberedImagePath) {
                $NumberedImages += @{
                    Path = $NumberedImagePath
                    Priority = 2
                    SortOrder = $i
                    Type = "Image$i"
                }
            }
        }
        
        $AllFoundImages += $NumberedImages | Sort-Object SortOrder
        
        # Priority 3: Lettered variants
        foreach ($letter in @('a','b','c','d','e','f')) {
            $LetterImagePath = Join-Path $ImageDir "$LotNumber-$letter.jpg"
            if (Test-Path $LetterImagePath) {
                $AllFoundImages += @{
                    Path = $LetterImagePath
                    Priority = 3
                    SortOrder = 100 + [int][char]$letter
                    Type = "Variant-$letter"
                }
            }
        }
        
        if ($AllFoundImages.Count -gt 0) {
            $SelectedImages = $AllFoundImages | 
                Sort-Object Priority, SortOrder | 
                Select-Object -First $MaxImagesPerLot
            
            $ImagePaths = $SelectedImages | ForEach-Object { $_.Path }
            
            Write-Log "Lot $LotNumber`: Found $($AllFoundImages.Count) images, including $($ImagePaths.Count) (max: $MaxImagesPerLot)" -Level "DEBUG"
            
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
        Create Date: 01/07/2025
        Version: 1.3.0
        Change Date: 10/10/2025
        Change Purpose: Rewritten using array join for bulletproof parsing
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
        
        # Use absolute path for report
        $ReportPath = Join-Path (Resolve-Path $OutputPath).Path ("AuctionLots_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".html")
        
        # Build HTML using array that we'll join - NO PARSING ERRORS!
        $htmlLines = @()
        $htmlLines += '<!DOCTYPE html>'
        $htmlLines += '<html>'
        $htmlLines += '<head>'
        $htmlLines += "    <title>Auction Lot Images - $AuctionName</title>"
        $htmlLines += '    <style>'
        $htmlLines += '        @media print {'
        $htmlLines += '            .page { page-break-after: always; }'
        $htmlLines += '            .no-print { display: none; }'
        $htmlLines += '        }'
        $htmlLines += '        body {'
        $htmlLines += '            font-family: ''Segoe UI'', Arial, sans-serif;'
        $htmlLines += '            margin: 20px;'
        $htmlLines += '            background: #f5f5f5;'
        $htmlLines += '        }'
        $htmlLines += '        .header {'
        $htmlLines += '            background: #2c3e50;'
        $htmlLines += '            color: white;'
        $htmlLines += '            padding: 20px;'
        $htmlLines += '            text-align: center;'
        $htmlLines += '            margin-bottom: 20px;'
        $htmlLines += '        }'
        $htmlLines += '        .page {'
        $htmlLines += '            background: white;'
        $htmlLines += '            padding: 20px;'
        $htmlLines += '            margin-bottom: 20px;'
        $htmlLines += '            box-shadow: 0 2px 5px rgba(0,0,0,0.1);'
        $htmlLines += '        }'
        $htmlLines += '        .lot-number {'
        $htmlLines += '            font-size: 24px;'
        $htmlLines += '            font-weight: bold;'
        $htmlLines += '            color: #2c3e50;'
        $htmlLines += '            margin-bottom: 10px;'
        $htmlLines += '        }'
        $htmlLines += '        .description {'
        $htmlLines += '            font-size: 14px;'
        $htmlLines += '            color: #666;'
        $htmlLines += '            margin-bottom: 15px;'
        $htmlLines += '        }'
        $htmlLines += '        .image-container {'
        $htmlLines += '            display: flex;'
        $htmlLines += '            flex-wrap: wrap;'
        $htmlLines += '            gap: 10px;'
        $htmlLines += '            margin-bottom: 10px;'
        $htmlLines += '        }'
        $htmlLines += '        .image-wrapper {'
        $htmlLines += '            flex: 1 1 48%;'
        $htmlLines += '            min-width: 300px;'
        $htmlLines += '        }'
        $htmlLines += '        img {'
        $htmlLines += '            width: 100%;'
        $htmlLines += '            height: auto;'
        $htmlLines += '            border: 1px solid #ddd;'
        $htmlLines += '            border-radius: 5px;'
        $htmlLines += '        }'
        $htmlLines += '        .image-label {'
        $htmlLines += '            text-align: center;'
        $htmlLines += '            font-size: 12px;'
        $htmlLines += '            color: #666;'
        $htmlLines += '            margin-top: 5px;'
        $htmlLines += '        }'
        $htmlLines += '        .print-btn {'
        $htmlLines += '            position: fixed;'
        $htmlLines += '            top: 20px;'
        $htmlLines += '            right: 20px;'
        $htmlLines += '            padding: 10px 20px;'
        $htmlLines += '            background: #3498db;'
        $htmlLines += '            color: white;'
        $htmlLines += '            border: none;'
        $htmlLines += '            border-radius: 5px;'
        $htmlLines += '            cursor: pointer;'
        $htmlLines += '            z-index: 1000;'
        $htmlLines += '        }'
        $htmlLines += '    </style>'
        $htmlLines += '</head>'
        $htmlLines += '<body>'
        $htmlLines += '    <button class="print-btn no-print" onclick="window.print()">Print to PDF</button>'
        $htmlLines += '    <div class="header">'
        $htmlLines += '        <h1>Auction Lot Images</h1>'
        $htmlLines += '        <div>'
        $htmlLines += "            Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')<br>"
        $htmlLines += "            Total Lots: $($Images.Count)<br>"
        $htmlLines += "            Total Images: $totalImageCount"
        $htmlLines += '        </div>'
        $htmlLines += '    </div>'
        
        # Add each lot
        foreach ($LotInfo in $Images) {
            $htmlLines += '    <div class="page">'
            $htmlLines += "        <div class=`"lot-number`">Lot #$($LotInfo.LotNumber)</div>"
            $htmlLines += "        <div class=`"description`">$($LotInfo.Description)</div>"
            $htmlLines += '        <div class="image-container">'
            
            if ($LotInfo.ImagePaths) {
                $imageNum = 1
                $totalForLot = $LotInfo.ImagePaths.Count
                
                foreach ($imagePath in $LotInfo.ImagePaths) {
                    $imageLabel = if ($totalForLot -gt 1) { 
                        "Image $imageNum of $totalForLot" 
                    } else { 
                        "Lot Image" 
                    }
                    
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
                        
                        $htmlLines += '            <div class="image-wrapper">'
                        $htmlLines += "                <img src=`"$imageSrc`" alt=`"Lot $($LotInfo.LotNumber) - $imageLabel`" />"
                        $htmlLines += "                <div class=`"image-label`">$imageLabel</div>"
                        $htmlLines += '            </div>'
                        
                        $imageNum++
                    }
                }
            }
            else {
                $htmlLines += '            <div class="image-wrapper">'
                $htmlLines += '                <div style="padding: 20px; background: #f0f0f0; text-align: center;">'
                $htmlLines += '                    No images available for this lot'
                $htmlLines += '                </div>'
                $htmlLines += '            </div>'
            }
            
            $htmlLines += '        </div>'
            $htmlLines += '    </div>'
        }
        
        # Close HTML
        $htmlLines += '</body>'
        $htmlLines += '</html>'
        
        # Join all lines with newlines and write to file
        $html = $htmlLines -join "`n"
        $html | Out-File -FilePath $ReportPath -Encoding UTF8
        
        Write-Log "Image report generated successfully: $ReportPath" -Level "SUCCESS"
        $script:ProcessingStats.PDFsGenerated++
        
        Start-Process $ReportPath
        
        return $ReportPath
    }
    catch {
        Write-Log "Failed to generate image report: $_" -Level "ERROR"
        throw
    }
}
#endregion

#region HTML Email Generation
function New-LogisticsEmailHTML {
    <#
    .SYNOPSIS
        Generates formatted HTML email using config file data when available
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 01/07/2025
        Version: 1.3.0
        Change Date: 10/10/2025
        Change Purpose: Rewritten using array join for bulletproof parsing
    #>
    param (
        [array]$LotData,
        [string]$PDFPath,
        [object]$Config = $null
    )
    
    try {
        Write-Log "Generating HTML email for $($LotData.Count) lots" -Level "INFO"
        
        # Get data from config if available, otherwise use defaults/CSV
        $pickupAddress = if ($Config -and $Config.auction_info.pickup_address) {
            $Config.auction_info.pickup_address
        } elseif ($LotData[0].Address) {
            $LotData[0].Address.Trim()
        } else {
            "[Pickup address not specified]"
        }
        
        $deliveryAddress = if ($Config -and $Config.delivery_address) {
            $Config.delivery_address
        } else {
            "1218 Lake Avenue, Ashtabula, OH 44004"
        }
        
        $logisticsContactPhone = if ($Config -and $Config.auction_info.logistics_contact.phone) {
            $Config.auction_info.logistics_contact.phone
        } else {
            "[Not available]"
        }
        
        $logisticsContactEmail = if ($Config -and $Config.auction_info.logistics_contact.email) {
            $Config.auction_info.logistics_contact.email
        } else {
            "[Not available]"
        }
        
        $pickupDateTime = if ($Config -and $Config.auction_info.pickup_datetime) {
            $Config.auction_info.pickup_datetime
        } else {
            "[To be determined]"
        }
        
        $deliveryDateTime = if ($Config -and $Config.auction_info.delivery_datetime) {
            $Config.auction_info.delivery_datetime
        } else {
            "[To be determined]"
        }
        
        # Build shipping requirements section
        $shippingReqsLines = @()
        if ($Config -and $Config.shipping_requirements) {
            $reqs = $Config.shipping_requirements
            $shippingReqsLines += '    <p><strong>SHIPPING REQUIREMENTS:</strong><br>'
            if ($reqs.truck_types) { $shippingReqsLines += "    Truck Types: $($reqs.truck_types)<br>" }
            if ($reqs.labor_needed) { $shippingReqsLines += "    Labor Needed: $($reqs.labor_needed)<br>" }
            if ($reqs.total_pallets) { $shippingReqsLines += "    Estimated Pallets: $($reqs.total_pallets)<br>" }
            if ($reqs.weight_notes) { $shippingReqsLines += "    Weight: $($reqs.weight_notes)<br>" }
            $shippingReqsLines += '    </p>'
        }
        $shippingReqs = $shippingReqsLines -join "`n"
        
        # Build special notes section
        $specialNotesLines = @()
        if ($Config -and $Config.auction_info.special_notes) {
            $specialNotesLines += '    <p><strong>SPECIAL NOTES:</strong></p>'
            $specialNotesLines += '    <ul>'
            foreach ($note in $Config.auction_info.special_notes) {
                $specialNotesLines += "        <li>$note</li>"
            }
            $specialNotesLines += '    </ul>'
        }
        $specialNotes = $specialNotesLines -join "`n"
        
        # Build items list with lot numbers
        $itemsLines = @()
        foreach ($Lot in $LotData) {
            $qtyText = if ($Lot.Quantity) { " (Qty: $($Lot.Quantity))" } else { "" }
            $itemsLines += "        <li><strong>Lot #$($Lot.Lot):</strong> $($Lot.Description)$qtyText</li>"
        }
        $ItemsList = $itemsLines -join "`n"
        
        # Build complete HTML using array
        $htmlLines = @()
        $htmlLines += '<!DOCTYPE html>'
        $htmlLines += '<html>'
        $htmlLines += '<head>'
        $htmlLines += '    <meta charset="UTF-8">'
        $htmlLines += '    <title>Logistics Quote Request</title>'
        $htmlLines += '    <style>'
        $htmlLines += '        body {'
        $htmlLines += '            font-family: Calibri, Arial, sans-serif;'
        $htmlLines += '            font-size: 11pt;'
        $htmlLines += '            line-height: 1.5;'
        $htmlLines += '            color: #000000;'
        $htmlLines += '            max-width: 700px;'
        $htmlLines += '            margin: 20px;'
        $htmlLines += '        }'
        $htmlLines += '        p { margin: 10px 0; }'
        $htmlLines += '        ul { margin-left: 30px; }'
        $htmlLines += '        li { margin: 5px 0; }'
        $htmlLines += '    </style>'
        $htmlLines += '</head>'
        $htmlLines += '<body>'
        $htmlLines += '    <p>Hello.</p>'
        $htmlLines += '    '
        $htmlLines += "    <p>We have an online auction that's closed with $($LotData.Count) lots we won to pick up. Here is the information:</p>"
        $htmlLines += '    '
        $htmlLines += "    <p><strong>PICKUP ADDRESS:</strong> $pickupAddress</p>"
        $htmlLines += '    '
        $htmlLines += "    <p><strong>DELIVERY ADDRESS:</strong> $deliveryAddress</p>"
        $htmlLines += '    '
        $htmlLines += '    <p><strong>AUCTION LOGISTICS CONTACT:</strong><br>'
        $htmlLines += "    Phone: $logisticsContactPhone<br>"
        $htmlLines += "    Email: $logisticsContactEmail</p>"
        $htmlLines += '    '
        $htmlLines += "    <p><strong>PICKUP DATE/TIME:</strong> $pickupDateTime</p>"
        $htmlLines += '    '
        $htmlLines += "    <p><strong>DELIVERY DATE/TIME:</strong> $deliveryDateTime</p>"
        $htmlLines += '    '
        if ($shippingReqs) {
            $htmlLines += $shippingReqs
            $htmlLines += '    '
        }
        if ($specialNotes) {
            $htmlLines += $specialNotes
            $htmlLines += '    '
        }
        $htmlLines += '    <p>The items (pictures in the attached PDF referenced by Lot Number) are:</p>'
        $htmlLines += '    '
        $htmlLines += '    <ul>'
        $htmlLines += $ItemsList
        $htmlLines += '    </ul>'
        $htmlLines += '    '
        $htmlLines += '    <p>Please send me a quote at your earliest opportunity.</p>'
        $htmlLines += '    '
        $htmlLines += '    <p>Thank you.</p>'
        $htmlLines += '    '
        $htmlLines += "    <p>John O'Neill Sr.<br>"
        $htmlLines += '    AWS Solutions LLC dba JT Custom Trailers<br>'
        $htmlLines += '    (440) 813-6695</p>'
        $htmlLines += '</body>'
        $htmlLines += '</html>'
        
        # Join all lines
        $HTML = $htmlLines -join "`n"
        
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
        Create Date: 01/07/2025
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
        
        $Outlook = New-Object -ComObject Outlook.Application
        $Mail = $Outlook.CreateItem(0)
        
        $Mail.Subject = $Subject
        $Mail.HTMLBody = $HTMLContent
        
        if ($To) {
            $Mail.To = $To
        }
        
        # Add attachments with absolute paths
        foreach ($Attachment in $Attachments) {
            if (Test-Path $Attachment) {
                # Convert to absolute path if relative
                $AbsolutePath = (Resolve-Path $Attachment).Path
                $Mail.Attachments.Add($AbsolutePath) | Out-Null
                Write-Log "Added attachment: $(Split-Path $AbsolutePath -Leaf)" -Level "DEBUG"
            }
            else {
                Write-Log "Attachment not found: $Attachment" -Level "WARNING"
            }
        }
        
        $Mail.Save()
        
        if ($Display) {
            $Mail.Display()
        }
        
        Write-Log "Outlook draft email created successfully" -Level "SUCCESS"
        
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
    
    Write-Host "║  📊 PROCESSING SUMMARY                                    ║" -ForegroundColor Yellow
    Write-Host "║  ─────────────────────────────────────────────────────    ║" -ForegroundColor Gray
    Write-Host "║  Total Lots Processed: $($Stats.TotalLots.ToString().PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║  Images Found:         $($Stats.ImagesFound.ToString().PadLeft(10))                          ║" -ForegroundColor Green
    Write-Host "║  Images Missing:       $($Stats.ImagesMissing.ToString().PadLeft(10))                          ║" -ForegroundColor $(if ($Stats.ImagesMissing -gt 0) { "Yellow" } else { "White" })
    Write-Host "║  PDFs Generated:       $($Stats.PDFsGenerated.ToString().PadLeft(10))                          ║" -ForegroundColor Green
    Write-Host "║  Emails Generated:     $($Stats.EmailsGenerated.ToString().PadLeft(10))                          ║" -ForegroundColor Green
    Write-Host "║                                                            ║" -ForegroundColor Cyan
    
    Write-Host "║  ⏱️  PERFORMANCE METRICS                                   ║" -ForegroundColor Yellow
    Write-Host "║  ─────────────────────────────────────────────────────    ║" -ForegroundColor Gray
    Write-Host "║  Processing Time:      $($Duration.ToString('mm\:ss').PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║  Start Time:           $(($Stats.StartTime.ToString('HH:mm:ss')).PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║  End Time:             $(($EndTime.ToString('HH:mm:ss')).PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║                                                            ║" -ForegroundColor Cyan
    
    $LotsWithImages = if ($Stats.TotalLots -gt 0 -and $Stats.ImagesMissing -ge 0) {
        $Stats.TotalLots - $Stats.ImagesMissing
    } else { $Stats.TotalLots }
    
    $SuccessRate = if ($Stats.TotalLots -gt 0) { 
        [math]::Round(($LotsWithImages / $Stats.TotalLots) * 100, 2) 
    } else { 0 }
    
    $SuccessRate = [math]::Min($SuccessRate, 100)
    
    Write-Host "║  📈 LOT IMAGE COVERAGE                                     ║" -ForegroundColor Yellow
    Write-Host "║  ─────────────────────────────────────────────────────    ║" -ForegroundColor Gray
    
    $BarLength = 40
    $FilledLength = [math]::Floor($SuccessRate / 100 * $BarLength)
    $FilledLength = [math]::Max(0, [math]::Min($FilledLength, $BarLength))
    $EmptyLength = $BarLength - $FilledLength
    $Bar = "█" * $FilledLength + "░" * $EmptyLength
    
    Write-Host "║  [$Bar] $($SuccessRate.ToString('F2'))%  ║" -ForegroundColor $(if ($SuccessRate -ge 80) { "Green" } elseif ($SuccessRate -ge 50) { "Yellow" } else { "Red" })
    Write-Host "║  Lots with images: $LotsWithImages of $($Stats.TotalLots)                             ║" -ForegroundColor White
    
    if ($Stats.ImagesFound -gt $Stats.TotalLots) {
        $avgImages = [math]::Round($Stats.ImagesFound / $LotsWithImages, 1)
        Write-Host "║  Average images per lot: $($avgImages.ToString('F1'))                             ║" -ForegroundColor White
    }
    
    Write-Host "║                                                            ║" -ForegroundColor Cyan
    
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
        Create Date: 01/07/2025
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
    # Determine log directory based on config file location if provided
    if ($ConfigPath -and (Test-Path $ConfigPath)) {
        $configDir = Split-Path $ConfigPath -Parent
        $LogDirectory = Join-Path $configDir "Logs"
        
        # Also set Output directory to be with the config if not explicitly specified
        if ($PSBoundParameters.ContainsKey('OutputDirectory') -eq $false) {
            $OutputDirectory = Join-Path $configDir "Output"
        }
        
        Write-Host "Using directories based on config location:" -ForegroundColor Cyan
        Write-Host "  Logs: $LogDirectory" -ForegroundColor Gray
        Write-Host "  Output: $OutputDirectory" -ForegroundColor Gray
    }
    
    Initialize-Logging -LogDir $LogDirectory
    
    if (!(Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
        Write-Log "Created output directory: $OutputDirectory" -Level "INFO"
    }
    
    # Load configuration if provided
    $config = $null
    if ($ConfigPath) {
        $config = Get-AuctionConfig -ConfigPath $ConfigPath
        if ($config) {
            Write-Log "Using configuration from: $ConfigPath" -Level "SUCCESS"
        }
    }
    
    # Import auction data
    Write-Log "Starting data import process" -Level "INFO"
    $AuctionData = Import-AuctionData -CSVPath $CSVPath
    
    # Process lot images
    Write-Log "Processing lot images (max $MaxImagesPerLot per lot)" -Level "INFO"
    $LotImages = Get-LotImages -Lots $AuctionData -ImageDir $ImageDirectory -MaxImagesPerLot $MaxImagesPerLot

    # Generate image report and convert to PDF
    $PDFFilePath = $null
    $ImageReportPath = $null

    if ($LotImages.Count -gt 0) {
        Write-Log "Generating HTML image report" -Level "INFO"
        $ImageReportPath = New-LotPDF -Images $LotImages -OutputPath $OutputDirectory -AuctionName "Auction"
        
        # Try automatic PDF conversion using helper script
        if ($script:PDFHelperAvailable) {
            Write-Log "Attempting automatic PDF conversion..." -Level "INFO"
            
            $PDFPath = Join-Path $OutputDirectory ("AuctionLots_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".pdf")
            
            Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
            Write-Host "║         PDF CONVERSION - DETAILED OUTPUT                  ║" -ForegroundColor Cyan
            Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan
            
            try {
                # Don't use -Quiet so we can see all debugging output
                $PDFFilePath = ConvertTo-PDFHelper -HTMLPath $ImageReportPath -OutputPath $PDFPath -Method "Auto"
                
                if ($PDFFilePath -and (Test-Path $PDFFilePath)) {
                    Write-Log "PDF created successfully: $(Split-Path $PDFFilePath -Leaf)" -Level "SUCCESS"
                    Write-Host "`n✓ PDF saved to: $PDFFilePath" -ForegroundColor Green
                    
                    # Open the folder containing the PDF
                    $pdfFolder = Split-Path $PDFFilePath -Parent
                    Start-Process explorer.exe -ArgumentList "/select,`"$PDFFilePath`""
                }
                else {
                    Write-Log "Automatic PDF conversion failed, using HTML report" -Level "WARNING"
                    Write-Host "`n⚠️  PDF conversion did not produce a file" -ForegroundColor Yellow
                    Write-Host "   Using HTML report instead: $ImageReportPath" -ForegroundColor Gray
                    $PDFFilePath = $ImageReportPath
                }
            }
            catch {
                Write-Log "Error during PDF conversion: $_" -Level "ERROR"
                Write-Log "Falling back to HTML report" -Level "WARNING"
                Write-Host "`n✗ Error: $_" -ForegroundColor Red
                $PDFFilePath = $ImageReportPath
            }
        }
        else {
            Write-Log "PDF helper not available, using HTML report" -Level "WARNING"
            Write-Host "`n" + ("=" * 70) -ForegroundColor Yellow
            Write-Host "MANUAL PDF CREATION:" -ForegroundColor Yellow
            Write-Host ("=" * 70) -ForegroundColor Yellow
            Write-Host "1. HTML report opened in your browser" -ForegroundColor White
            Write-Host "2. Press Ctrl+P to print" -ForegroundColor White
            Write-Host "3. Select 'Save as PDF' or your Foxit PDF Printer" -ForegroundColor White
            Write-Host "4. Save to Output folder" -ForegroundColor White
            Write-Host ("=" * 70) + "`n" -ForegroundColor Yellow
            $PDFFilePath = $ImageReportPath
        }
    }
    else {
        Write-Log "No images found, skipping PDF generation" -Level "WARNING"
    }
    
    # Generate HTML email using config data
    Write-Log "Generating HTML email with config data" -Level "INFO"
    $EmailHTML = New-LogisticsEmailHTML -LotData $AuctionData -PDFPath $PDFFilePath -Config $config
    
    # Determine email subject
    $emailSubject = if ($config -and $config.email_subject) {
        $config.email_subject
    } else {
        "Freight Quote Request - $(Get-Date -Format 'yyyy-MM-dd')"
    }
    
    # Save HTML to file with absolute path
    $OutputDirAbsolute = if (Test-Path $OutputDirectory) {
        (Resolve-Path $OutputDirectory).Path
    } else {
        (New-Item -ItemType Directory -Path $OutputDirectory -Force).FullName
    }
    
    $HTMLFilePath = Join-Path $OutputDirAbsolute ("LogisticsEmail_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".html")
    $EmailHTML | Out-File -FilePath $HTMLFilePath -Encoding UTF8
    Write-Log "HTML email saved to: $HTMLFilePath" -Level "SUCCESS"
    Write-Host "`n================================================" -ForegroundColor Cyan
    Write-Host "EMAIL FILE LOCATION:" -ForegroundColor Yellow
    Write-Host $HTMLFilePath -ForegroundColor Green
    Write-Host "================================================`n" -ForegroundColor Cyan

    # Prepare attachments
    $AttachmentList = @()
    if ($PDFFilePath) {
        $AttachmentList += $PDFFilePath
    }

    # Create Outlook draft
    if ($CreateOutlookDraft) {
        $OutlookCreated = New-OutlookDraftEmail -HTMLContent $EmailHTML -Subject $emailSubject -Attachments $AttachmentList -Display
    }
    else {
        $OutlookCreated = $false
        Write-Log "Outlook draft creation skipped (use -CreateOutlookDraft to enable)" -Level "INFO"
    }

    if (-not $OutlookCreated) {
        Write-Log "Opening HTML in default browser" -Level "INFO"
        Start-Process $HTMLFilePath
    }

    # Generate processing report
    $ReportPath = $null
    try {
        $ReportPath = Export-ProcessingReport -OutputPath $OutputDirectory -Stats $script:ProcessingStats -LotData $AuctionData
    }
    catch {
        Write-Log "Skipping processing report: $($_.Exception.Message)" -Level "WARNING"
    }
    
    $script:ProcessingStats.ProcessedLots = $AuctionData.Count
    
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