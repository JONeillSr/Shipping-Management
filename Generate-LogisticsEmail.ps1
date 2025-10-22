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
    Version: 1.5.5
    Change Date: 10/17/2025
    Change Purpose: Fixed all HTML parsing errors using proper here-strings

.CHANGELOG
    1.5.5 - 10/17/2025 - Fixed HTML parsing errors in New-LotPDF and New-LogisticsEmailHTML
                       - All HTML generation now uses proper here-strings
                       - Resolved '<' operator and onclick parsing issues
                       - Improved error messages and logging
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
        Write-Host "   PDF conversion will not be available" -ForegroundColor Yellow
    }
}
else {
    Write-Host "⚠️  PDF helper not found: $helperScript" -ForegroundColor Yellow
    Write-Host "   Download from: https://github.com/JONeillSr/Shipping-Management" -ForegroundColor Cyan
    Write-Host "   Continuing without automatic PDF conversion..." -ForegroundColor Yellow
}
#endregion

#region Module Requirements
# Check and install required modules
$RequiredModules = @('PSWritePDF', 'ImportExcel')
foreach ($Module in $RequiredModules) {
    if (!(Get-Module -ListAvailable -Name $Module)) {
        Write-Host "Installing module: $Module" -ForegroundColor Cyan
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

    Write-JTLSLog "=== Logistics Email Automation Started ===" -Level "INFO"
    Write-JTLSLog "Script Version: 1.5.5" -Level "INFO"
    Write-JTLSLog "User: $env:USERNAME" -Level "INFO"
    Write-JTLSLog "Machine: $env:COMPUTERNAME" -Level "INFO"
    Write-JTLSLog "PDF Helper Available: $script:PDFHelperAvailable" -Level "INFO"
}

function Write-JTLSLog {
    <#
    .SYNOPSIS
        Writes detailed log entries
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 01/07/2025
        Version: 1.0.0
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
        default   { Write-Host $LogEntry -ForegroundColor White }
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
    #>
    param(
        [string]$ConfigPath
    )

    if (-not $ConfigPath -or -not (Test-Path $ConfigPath)) {
        Write-JTLSLog "No config file provided or file not found" -Level "DEBUG"
        return $null
    }

    try {
        Write-JTLSLog "Loading configuration from: $ConfigPath" -Level "INFO"
        $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        Write-JTLSLog "Configuration loaded successfully" -Level "SUCCESS"
        return $config
    }
    catch {
        Write-JTLSLog "Failed to load configuration: $_" -Level "ERROR"
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
    #>
    param (
        [string]$CSVPath
    )

    try {
        Write-JTLSLog "Importing CSV data from: $CSVPath" -Level "INFO"
        $script:ProcessingStats.DataSources += $CSVPath

        $Data = Import-Csv -Path $CSVPath
        $script:ProcessingStats.TotalLots = $Data.Count

        Write-JTLSLog "Successfully imported $($Data.Count) lots" -Level "SUCCESS"

        # Validate required columns
        $RequiredColumns = @('Lot', 'Description')
        $MissingColumns = $RequiredColumns | Where-Object { $_ -notin $Data[0].PSObject.Properties.Name }

        if ($MissingColumns) {
            throw "Missing required columns: $($MissingColumns -join ', ')"
        }

        return $Data
    }
    catch {
        Write-JTLSLog "Failed to import CSV: $_" -Level "ERROR"
        throw
    }
}

function Get-LotImage {
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

            Write-JTLSLog "Lot $LotNumber`: Found $($AllFoundImages.Count) images, including $($ImagePaths.Count) (max: $MaxImagesPerLot)" -Level "DEBUG"

            $selectedNames = $SelectedImages | ForEach-Object { Split-Path $_.Path -Leaf }
            Write-JTLSLog "  Selected: $($selectedNames -join ', ')" -Level "DEBUG"

            if ($AllFoundImages.Count -gt $MaxImagesPerLot) {
                Write-JTLSLog "Lot $LotNumber has $($AllFoundImages.Count) images, selected: $(($SelectedImages.Type) -join ', ')" -Level "INFO"
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
            Write-JTLSLog "No images found for Lot $LotNumber" -Level "WARNING"
            $script:ProcessingStats.ImagesMissing++
        }
    }

    $totalIncluded = ($ImageList.ImageCount | Measure-Object -Sum).Sum
    $totalFound = ($ImageList.TotalFound | Measure-Object -Sum).Sum
    Write-JTLSLog "Including $totalIncluded of $totalFound total images found" -Level "INFO"

    return $ImageList
}
#endregion

#region PDF Generation
function New-LotPDF {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    param (
        [array]$Images,
        [string]$OutputPath,
        [string]$AuctionName
    )

    try {
        $totalImageCount = 0
        foreach ($lot in $Images) {
            if ($lot.ImagePaths) {
                $totalImageCount += $lot.ImagePaths.Count
            }
        }

        Write-JTLSLog "Generating image report for $($Images.Count) lots with $totalImageCount total images" -Level "INFO"

        $ReportPath = Join-Path (Resolve-Path $OutputPath).Path ("AuctionLots_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".html")

        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        $writer = New-Object System.IO.StreamWriter($ReportPath, $false, $utf8NoBom)

        try {
            $writer.WriteLine('<!DOCTYPE html>')
            $writer.WriteLine('<html>')
            $writer.WriteLine('<head>')
            $writer.WriteLine('    <meta charset="UTF-8">')
            $writer.WriteLine('    <title>Auction Lot Images - ' + $AuctionName + '</title>')
            $writer.WriteLine('    <style>')
            $writer.WriteLine('        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }')
            $writer.WriteLine('        .header { background: #2c3e50; color: white; padding: 20px; text-align: center; margin-bottom: 20px; border-radius: 8px; }')
            $writer.WriteLine('        .page { background: white; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); border-radius: 8px; page-break-after: always; }')
            $writer.WriteLine('        .lot-number { font-size: 24px; font-weight: bold; color: #2c3e50; margin-bottom: 10px; border-bottom: 2px solid #3498db; padding-bottom: 5px; }')
            $writer.WriteLine('        .description { font-size: 14px; color: #666; margin-bottom: 15px; font-style: italic; }')
            $writer.WriteLine('        .image-container { display: flex; flex-wrap: wrap; gap: 15px; margin-bottom: 10px; }')
            $writer.WriteLine('        .image-wrapper { flex: 1 1 48%; min-width: 300px; background: #f8f9fa; padding: 10px; border-radius: 5px; }')
            $writer.WriteLine('        img { width: 100%; height: auto; border: 2px solid #ddd; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }')
            $writer.WriteLine('        .image-label { text-align: center; font-size: 12px; color: #666; margin-top: 8px; font-weight: 500; }')
            $writer.WriteLine('        .print-btn { position: fixed; top: 20px; right: 20px; padding: 12px 24px; background: #3498db; color: white; border: none; border-radius: 5px; cursor: pointer; z-index: 1000; font-size: 14px; font-weight: bold; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }')
            $writer.WriteLine('        .print-btn:hover { background: #2980b9; }')
            $writer.WriteLine('        .stats { margin-top: 10px; font-size: 14px; }')
            $writer.WriteLine('        @media print { .page { page-break-after: always; } .no-print { display: none; } }')
            $writer.WriteLine('    </style>')
            $writer.WriteLine('</head>')
            $writer.WriteLine('<body>')

            # Use character codes for problematic syntax
            $btnHtml = '    <button class="print-btn no-print" onclick="window.print' + [char]40 + [char]41 + '">Print to PDF</button>'
            $writer.WriteLine($btnHtml)

            $writer.WriteLine('    <div class="header">')
            $writer.WriteLine('        <h1>Auction Lot Images</h1>')
            $writer.WriteLine('        <div class="stats">')
            $writer.WriteLine('            Generated: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm') + '<br>')
            $writer.WriteLine('            Total Lots: ' + $Images.Count + '<br>')
            $writer.WriteLine('            Total Images: ' + $totalImageCount)
            $writer.WriteLine('        </div>')
            $writer.WriteLine('    </div>')

            foreach ($LotInfo in $Images) {
                $writer.WriteLine('    <div class="page">')
                $writer.WriteLine('        <div class="lot-number">Lot #' + $LotInfo.LotNumber + '</div>')
                $writer.WriteLine('        <div class="description">' + $LotInfo.Description + '</div>')
                $writer.WriteLine('        <div class="image-container">')

                if ($LotInfo.ImagePaths -and $LotInfo.ImagePaths.Count -gt 0) {
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
                                $imageSrc = 'data:image/jpeg;base64,' + $imageBase64

                                $writer.WriteLine('            <div class="image-wrapper">')
                                $writer.WriteLine('                <img src="' + $imageSrc + '" alt="Lot ' + $LotInfo.LotNumber + ' - ' + $imageLabel + '" />')
                                $writer.WriteLine('                <div class="image-label">' + $imageLabel + '</div>')
                                $writer.WriteLine('            </div>')

                                $imageNum++
                            }
                            catch {
                                Write-JTLSLog "Could not embed image: $imagePath - $_" -Level "WARNING"
                            }
                        }
                    }
                }
                else {
                    $writer.WriteLine('            <div class="image-wrapper">')
                    $writer.WriteLine('                <div style="padding: 20px; background: #f0f0f0; text-align: center; border-radius: 5px;">')
                    $writer.WriteLine('                    <p style="color: #666; margin: 0;">No images available for this lot</p>')
                    $writer.WriteLine('                </div>')
                    $writer.WriteLine('            </div>')
                }

                $writer.WriteLine('        </div>')
                $writer.WriteLine('    </div>')
            }

            $writer.WriteLine('</body>')
            $writer.WriteLine('</html>')
        }
        finally {
            if ($writer) {
                $writer.Close()
                $writer.Dispose()
            }
        }

        Write-JTLSLog "Image report generated successfully: $ReportPath" -Level "SUCCESS"
        Write-JTLSLog "  Lots: $($Images.Count)" -Level "INFO"
        Write-JTLSLog "  Images: $totalImageCount" -Level "INFO"

        $script:ProcessingStats.PDFsGenerated++
        Start-Process $ReportPath

        return $ReportPath
    }
    catch {
        Write-JTLSLog "Failed to generate image report: $_" -Level "ERROR"
        Write-JTLSLog "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
        throw
    }
}

#endregion

#region HTML Email Generation
function New-LogisticsEmailHTML {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    <#
    .SYNOPSIS
        Generates formatted HTML email using config file data when available
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 01/07/2025
        Version: 1.4.3
        Change Date: 10/17/2025
        Change Purpose: Using StringBuilder to avoid all parsing issues
    #>
    param (
        [array]$LotData,
        [string]$PDFPath,
        [object]$Config = $null
    )

    try {
        Write-JTLSLog "Generating HTML email for $($LotData.Count) lots" -Level "INFO"

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

        # Use StringBuilder for efficient string building
        $sb = New-Object System.Text.StringBuilder

        # Build HTML
        [void]$sb.AppendLine('<!DOCTYPE html>')
        [void]$sb.AppendLine('<html>')
        [void]$sb.AppendLine('<head>')
        [void]$sb.AppendLine('    <meta charset="UTF-8">')
        [void]$sb.AppendLine('    <title>Logistics Quote Request</title>')
        [void]$sb.AppendLine('    <style>')
        [void]$sb.AppendLine('        body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.5; color: #000000; max-width: 700px; margin: 20px; }')
        [void]$sb.AppendLine('        p { margin: 10px 0; }')
        [void]$sb.AppendLine('        ul { margin-left: 30px; }')
        [void]$sb.AppendLine('        li { margin: 5px 0; }')
        [void]$sb.AppendLine('        strong { color: #2c3e50; }')
        [void]$sb.AppendLine('    </style>')
        [void]$sb.AppendLine('</head>')
        [void]$sb.AppendLine('<body>')
        [void]$sb.AppendLine('    <p>Hello,</p>')
        [void]$sb.AppendLine('    ')
        [void]$sb.AppendLine('    <p>We have an online auction that' + [char]39 + 's closed with ' + $LotData.Count + ' lots we won to pick up. Here is the information:</p>')
        [void]$sb.AppendLine('    ')
        [void]$sb.AppendLine('    <p><strong>PICKUP ADDRESS:</strong> ' + $pickupAddress + '</p>')
        [void]$sb.AppendLine('    ')
        [void]$sb.AppendLine('    <p><strong>DELIVERY ADDRESS:</strong> ' + $deliveryAddress + '</p>')
        [void]$sb.AppendLine('    ')
        [void]$sb.AppendLine('    <p><strong>AUCTION LOGISTICS CONTACT:</strong><br>')
        [void]$sb.AppendLine('    Phone: ' + $logisticsContactPhone + '<br>')
        [void]$sb.AppendLine('    Email: ' + $logisticsContactEmail + '</p>')
        [void]$sb.AppendLine('    ')
        [void]$sb.AppendLine('    <p><strong>PICKUP DATE/TIME:</strong> ' + $pickupDateTime + '</p>')
        [void]$sb.AppendLine('    ')
        [void]$sb.AppendLine('    <p><strong>DELIVERY DATE/TIME:</strong> ' + $deliveryDateTime + '</p>')
        [void]$sb.AppendLine('    ')

        # Add shipping requirements if available
        if ($Config -and $Config.shipping_requirements) {
            $reqs = $Config.shipping_requirements
            [void]$sb.AppendLine('    <p><strong>SHIPPING REQUIREMENTS:</strong><br>')
            if ($reqs.truck_types) { [void]$sb.AppendLine('    Truck Types: ' + $reqs.truck_types + '<br>') }
            if ($reqs.labor_needed) { [void]$sb.AppendLine('    Labor Needed: ' + $reqs.labor_needed + '<br>') }
            if ($reqs.total_pallets) { [void]$sb.AppendLine('    Estimated Pallets: ' + $reqs.total_pallets + '<br>') }
            if ($reqs.weight_notes) { [void]$sb.AppendLine('    Weight: ' + $reqs.weight_notes + '<br>') }
            [void]$sb.AppendLine('    </p>')
            [void]$sb.AppendLine('    ')
        }

        # Add special notes if available
        if ($Config -and $Config.auction_info.special_notes) {
            [void]$sb.AppendLine('    <p><strong>SPECIAL NOTES:</strong></p>')
            [void]$sb.AppendLine('    <ul>')
            foreach ($note in $Config.auction_info.special_notes) {
                [void]$sb.AppendLine('        <li>' + $note + '</li>')
            }
            [void]$sb.AppendLine('    </ul>')
            [void]$sb.AppendLine('    ')
        }

        [void]$sb.AppendLine('    <p>The items (pictures in the attached PDF referenced by Lot Number) are:</p>')
        [void]$sb.AppendLine('    ')
        [void]$sb.AppendLine('    <ul>')

        # Add items list
        foreach ($Lot in $LotData) {
            $qtyText = if ($Lot.Quantity) { ' (Qty: ' + $Lot.Quantity + ')' } else { '' }
            [void]$sb.AppendLine('        <li><strong>Lot #' + $Lot.Lot + ':</strong> ' + $Lot.Description + $qtyText + '</li>')
        }

        [void]$sb.AppendLine('    </ul>')
        [void]$sb.AppendLine('    ')
        [void]$sb.AppendLine('    <p>Please send me a quote at your earliest opportunity.</p>')
        [void]$sb.AppendLine('    ')
        [void]$sb.AppendLine('    <p>Thank you.</p>')
        [void]$sb.AppendLine('    ')
        [void]$sb.AppendLine('    <p>John O' + [char]39 + 'Neill Sr.<br>')
        [void]$sb.AppendLine('    AWS Solutions LLC dba JT Custom Trailers<br>')
        [void]$sb.AppendLine('    (440) 813-6695</p>')
        [void]$sb.AppendLine('</body>')
        [void]$sb.AppendLine('</html>')

        $HTML = $sb.ToString()

        Write-JTLSLog "HTML email generated successfully" -Level "SUCCESS"
        $script:ProcessingStats.EmailsGenerated++

        return $HTML
    }
    catch {
        Write-JTLSLog "Failed to generate HTML email: $_" -Level "ERROR"
        Write-JTLSLog "Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
        throw
    }
}

function New-OutlookDraftEmail {
    [CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
    <#
    .SYNOPSIS
        Creates a draft email in Outlook with HTML formatting
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 01/07/2025
        Version: 1.0.0
    #>
    param (
        [string]$HTMLContent,
        [string]$Subject = "Freight Quote Request - $(Get-Date -Format 'yyyy-MM-dd')",
        [string[]]$Attachments = @(),
        [string]$To = "",
        [switch]$Display
    )

    try {
        Write-JTLSLog "Creating Outlook draft email" -Level "INFO"

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
                Write-JTLSLog "Added attachment: $(Split-Path $AbsolutePath -Leaf)" -Level "DEBUG"
            }
            else {
                Write-JTLSLog "Attachment not found: $Attachment" -Level "WARNING"
            }
        }

        $Mail.Save()

        if ($Display) {
            $Mail.Display()
        }

        Write-JTLSLog "Outlook draft email created successfully" -Level "SUCCESS"

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Mail) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        return $true
    }
    catch {
        Write-JTLSLog "Failed to create Outlook email: $_" -Level "ERROR"
        Write-JTLSLog "Make sure Outlook is installed and configured" -Level "WARNING"
        return $false
    }
}
#endregion

#region Dashboard and Reporting
function Show-Dashboard {
    <#
    .SYNOPSIS
        Displays interactive processing dashboard
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 01/07/2025
        Version: 1.0.0
    #>
    param (
        [hashtable]$Stats
    )

    Write-Host "`nPress Enter to view dashboard..." -ForegroundColor Cyan
    Read-Host

    $EndTime = Get-Date
    $Duration = $EndTime - $Stats.StartTime

    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║         LOGISTICS EMAIL AUTOMATION DASHBOARD              ║" -ForegroundColor Cyan
    Write-Host "╠════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan
    Write-Host "║                                                            ║" -ForegroundColor Cyan

    Write-Host "║  📊 PROCESSING SUMMARY                                    ║" -ForegroundColor White
    Write-Host "║  ─────────────────────────────────────────────────────    ║" -ForegroundColor Gray
    Write-Host "║  Total Lots Processed: $($Stats.TotalLots.ToString().PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║  Images Found:         $($Stats.ImagesFound.ToString().PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║  Images Missing:       $($Stats.ImagesMissing.ToString().PadLeft(10))                          ║" -ForegroundColor $(if ($Stats.ImagesMissing -gt 0) { "Yellow" } else { "White" })
    Write-Host "║  PDFs Generated:       $($Stats.PDFsGenerated.ToString().PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║  Emails Generated:     $($Stats.EmailsGenerated.ToString().PadLeft(10))                          ║" -ForegroundColor White
    Write-Host "║                                                            ║" -ForegroundColor Cyan

    Write-Host "║  ⏱️  PERFORMANCE METRICS                                   ║" -ForegroundColor White
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

    Write-Host "║  📈 LOT IMAGE COVERAGE                                     ║" -ForegroundColor White
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

    Write-Host "║  📁 DATA SOURCES                                           ║" -ForegroundColor White
    Write-Host "║  ─────────────────────────────────────────────────────    ║" -ForegroundColor Gray
    foreach ($Source in $Stats.DataSources) {
        $SourceName = Split-Path $Source -Leaf
        if ($SourceName.Length -gt 50) {
            $SourceName = $SourceName.Substring(0, 47) + "..."
        }
        Write-Host "║  • $($SourceName.PadRight(54))  ║" -ForegroundColor Gray
    }

    Write-Host "║                                                            ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""

    if ($Stats.ImagesMissing -gt 0) {
        Write-Host "⚠️  WARNING: $($Stats.ImagesMissing) lots are missing images!" -ForegroundColor Yellow
        Write-Host "   Check the image directory for missing .jpg files" -ForegroundColor Yellow
        Write-Host ""
    }

    Write-Host "✅ Processing Complete!" -ForegroundColor Green
    Write-Host "   Logs saved to: $script:LogFile" -ForegroundColor Cyan
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
    Write-JTLSLog "Processing report exported to: $ReportPath" -Level "SUCCESS"

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
        Write-JTLSLog "Created output directory: $OutputDirectory" -Level "INFO"
    }

    # Load configuration if provided
    $config = $null
    if ($ConfigPath) {
        $config = Get-AuctionConfig -ConfigPath $ConfigPath
        if ($config) {
            Write-JTLSLog "Using configuration from: $ConfigPath" -Level "SUCCESS"
        }
    }

    # Import auction data
    Write-JTLSLog "Starting data import process" -Level "INFO"
    $AuctionData = Import-AuctionData -CSVPath $CSVPath

    # Process lot images
    Write-JTLSLog "Processing lot images (max $MaxImagesPerLot per lot)" -Level "INFO"
    $LotImages = Get-LotImage -Lots $AuctionData -ImageDir $ImageDirectory -MaxImagesPerLot $MaxImagesPerLot

    # Generate image report and convert to PDF
    $PDFFilePath = $null
    $ImageReportPath = $null

    if ($LotImages.Count -gt 0) {
        Write-JTLSLog "Generating HTML image report" -Level "INFO"
        $ImageReportPath = New-LotPDF -Images $LotImages -OutputPath $OutputDirectory -AuctionName "Auction"

        # Try automatic PDF conversion using helper script
        if ($script:PDFHelperAvailable) {
            Write-JTLSLog "Attempting automatic PDF conversion..." -Level "INFO"

            $PDFPath = Join-Path $OutputDirectory ("AuctionLots_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".pdf")

            Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
            Write-Host "║         PDF CONVERSION - DETAILED OUTPUT                  ║" -ForegroundColor Cyan
            Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

            try {
                # Don't use -Quiet so we can see all debugging output
                $PDFFilePath = ConvertTo-PDFHelper -HTMLPath $ImageReportPath -OutputPath $PDFPath -Method "Auto"

                if ($PDFFilePath -and (Test-Path $PDFFilePath)) {
                    Write-JTLSLog "PDF created successfully: $(Split-Path $PDFFilePath -Leaf)" -Level "SUCCESS"
                    Write-Host "`n✓ PDF saved to: $PDFFilePath" -ForegroundColor Green

                    # Open the folder containing the PDF
                    $pdfFolder = Split-Path $PDFFilePath -Parent
                    Start-Process explorer.exe -ArgumentList "/select,`"$PDFFilePath`""
                }
                else {
                    Write-JTLSLog "Automatic PDF conversion failed, using HTML report" -Level "WARNING"
                    Write-Host "`n⚠️  PDF conversion did not produce a file" -ForegroundColor Yellow
                    Write-Host "   Using HTML report instead: $ImageReportPath" -ForegroundColor Yellow
                    $PDFFilePath = $ImageReportPath
                }
            }
            catch {
                Write-JTLSLog "Error during PDF conversion: $_" -Level "ERROR"
                Write-JTLSLog "Falling back to HTML report" -Level "WARNING"
                Write-Host "`n✗ Error: $_" -ForegroundColor Red
                $PDFFilePath = $ImageReportPath
            }
        }
        else {
            Write-JTLSLog "PDF helper not available, using HTML report" -Level "WARNING"
            Write-Host "`n" + ("=" * 70) -ForegroundColor Cyan
            Write-Host "MANUAL PDF CREATION:" -ForegroundColor Yellow
            Write-Host ("=" * 70) -ForegroundColor Cyan
            Write-Host "1. HTML report opened in your browser" -ForegroundColor White
            Write-Host "2. Press Ctrl+P to print" -ForegroundColor White
            Write-Host "3. Select 'Save as PDF' or your Foxit PDF Printer" -ForegroundColor White
            Write-Host "4. Save to Output folder" -ForegroundColor White
            Write-Host ("=" * 70) + "`n" -ForegroundColor Cyan
            $PDFFilePath = $ImageReportPath
        }
    }
    else {
        Write-JTLSLog "No images found, skipping PDF generation" -Level "WARNING"
    }

    # Generate HTML email using config data
    Write-JTLSLog "Generating HTML email with config data" -Level "INFO"
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
    Write-JTLSLog "HTML email saved to: $HTMLFilePath" -Level "SUCCESS"
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
        Write-JTLSLog "Outlook draft creation skipped (use -CreateOutlookDraft to enable)" -Level "INFO"
    }

    if (-not $OutlookCreated) {
        Write-JTLSLog "Opening HTML in default browser" -Level "INFO"
        Start-Process $HTMLFilePath
    }

    # Generate processing report
    $ReportPath = $null
    try {
        $ReportPath = Export-ProcessingReport -OutputPath $OutputDirectory -Stats $script:ProcessingStats -LotData $AuctionData
    }
    catch {
        Write-JTLSLog "Skipping processing report: $($_.Exception.Message)" -Level "WARNING"
    }

    $script:ProcessingStats.ProcessedLots = $AuctionData.Count

    if ($ShowDashboard) {
        Show-Dashboard -Stats $script:ProcessingStats
    }

    Write-JTLSLog "=== Logistics Email Automation Completed Successfully ===" -Level "SUCCESS"
    Write-JTLSLog "Output files saved to: $OutputDirectory" -Level "INFO"

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
    Write-JTLSLog "Fatal error: $_" -Level "ERROR"
    Write-JTLSLog "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"

    [PSCustomObject]@{
        Success = $false
        Error = $_.Exception.Message
        LogFile = $script:LogFile
        ErrorLogFile = $script:ErrorLogFile
    }

    throw
}
finally {
    Write-JTLSLog "Script execution completed" -Level "INFO"
    Write-JTLSLog "Total execution time: $((Get-Date) - $script:ProcessingStats.StartTime)" -Level "INFO"
}
#endregion