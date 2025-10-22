<#
.SYNOPSIS
    Converts parsed invoice data to PurchaseTracking CSV format

.DESCRIPTION
    This script uses the Generic-PDF-Invoice-Parser.ps1 to parse an invoice PDF,
    then transforms the output into a CSV matching the PurchaseTracking format with
    columns for Lot, Description, Address, Plant, Location, Quantity, Bid, Sale Price,
    Premium, Tax, Rigging Fee, Freight Cost, Other Costs, Total, Per Item Cost, and
    additional tracking fields.

.PARAMETER PDFPath
    Path to the PDF invoice file to parse

.PARAMETER ParserScriptPath
    Path to the Generic-PDF-Invoice-Parser.ps1 script (default: .\Generic-PDF-Invoice-Parser.ps1)

.PARAMETER OutputCSVPath
    Path for the output CSV file (default: .\PurchaseTracking_YYYYMMDD_HHMMSS.csv)

.PARAMETER DefaultPlant
    Default plant number if not specified in invoice (default: blank)

.PARAMETER DefaultQuantity
    Default quantity per lot if not specified (default: 1)

.EXAMPLE
    .\Invoice-To-PurchaseTracking-CSV.ps1 -PDFPath ".\invoice.pdf"

.EXAMPLE
    .\Invoice-To-PurchaseTracking-CSV.ps1 -PDFPath ".\invoice.pdf" -OutputCSVPath ".\MyTracking.csv"

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 10/18/2025
    Version: 1.1.0
    Change Date: 10/18/2025
    Change Purpose: Enhanced HammerPrice extraction to handle both hashtable and PSCustomObject formats

.CHANGELOG
    v1.1.0: Fixed HammerPrice extraction; added proper handling for hashtable vs PSCustomObject
    v1.0.0: Initial release - Parse invoice and create PurchaseTracking CSV
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]$PDFPath,

    [Parameter(Mandatory=$false)]
    [string]$ParserScriptPath = ".\Generic-PDF-Invoice-Parser.ps1",

    [Parameter(Mandatory=$false)]
    [string]$OutputCSVPath,

    [Parameter(Mandatory=$false)]
    [string]$DefaultPlant = "",

    [Parameter(Mandatory=$false)]
    [int]$DefaultQuantity = 1
)

# Verify the parser script exists
if (-not (Test-Path $ParserScriptPath)) {
    Write-Error "Parser script not found at: $ParserScriptPath"
    Write-Output "Please ensure Generic-PDF-Invoice-Parser.ps1 is in the same directory or specify -ParserScriptPath"
    exit 1
}

# Verify the PDF exists
if (-not (Test-Path $PDFPath)) {
    Write-Error "PDF file not found at: $PDFPath"
    exit 1
}

Write-Output "`n╔════════════════════════════════════════════════════════╗"
Write-Output "║   INVOICE TO PURCHASE TRACKING CSV CONVERTER v1.0.0   ║"
Write-Output "╚════════════════════════════════════════════════════════╝`n"

Write-Output "📄 Processing invoice: $(Split-Path $PDFPath -Leaf)"

# Call the parser script with -ReturnObject to get the parsed data
Write-Output "`n🔄 Parsing invoice data..."
try {
    $parsedData = & $ParserScriptPath -PDFPath $PDFPath -ReturnObject

    # Handle case where the parser returns multiple objects or stray output
    if ($parsedData -is [array]) {
        $parsedData = $parsedData | Where-Object { $_ -isnot [string] } | Select-Object -Last 1
    }

    # Verify we actually got a structured object
    if (-not $parsedData -or -not ($parsedData.PSObject.Properties.Name -contains 'Items')) {
        Write-Error "Parser did not return a structured object with an 'Items' property."
        exit 1
    }

    Write-Output "   ✅ Successfully parsed invoice"

} catch {
    Write-Error "Error calling parser script: $_"
    exit 1
}

# Extract key information
$vendor = $parsedData.Vendor
$location = if ($parsedData.PickupAddresses.Count -gt 0) {
    $parsedData.PickupAddresses[0].OneLine
} else {
    ""
}

$address = if ($parsedData.PickupAddresses.Count -gt 0) {
    $parsedData.PickupAddresses[0].Street
} else {
    ""
}

# Create CSV rows for each item
Write-Output "   📦 Processing $($parsedData.Items.Count) items..."
$csvData = @()

foreach ($item in $parsedData.Items) {
    # Extract lot number and description
    $lotNumber = $item.LotNumber
    $description = $item.Description

    # Get hammer price - handle both hashtable and PSCustomObject
    $hammerPrice = ''
    if ($item -is [hashtable]) {
        if ($item.ContainsKey('HammerPrice')) {
            $hammerPrice = $item['HammerPrice']
        }
    } elseif ($item.HammerPrice) {
        $hammerPrice = $item.HammerPrice
    }

    # Create row object matching PurchaseTracking format
    $row = [PSCustomObject]@{
        'Lot'              = $lotNumber
        'Description'      = $description
        'Address'          = $address
        'Plant'            = $DefaultPlant
        'Location'         = $location
        'Quantity'         = $DefaultQuantity
        'Bid'              = ''  # Not in invoice
        'Sale Price'       = $hammerPrice  # Hammer Price = Sale Price
        'Premium'          = ''  # Calculated per item if needed
        'Tax'              = ''  # Individual tax not in invoice
        'Rigging Fee'      = ''  # Not in this invoice type
        'Freight Cost'     = ''  # Not in this invoice type
        'Other Costs'      = ''  # Not in this invoice type
        'Total'            = ''  # Per-item total not available
        'Per Item Cost'    = $hammerPrice  # Same as sale price for now
        'Category'         = ''  # User fills in later
        'Notes'            = ''  # User fills in later
        'Photos'           = ''  # User fills in later
        'MSRP'             = ''  # User research later
        'Average On-Line'  = ''  # User research later
        'Our Asking Price' = ''  # User fills in later
        'Sold Price'       = ''  # User fills when sold
    }

    $csvData += $row
}

# Generate output filename if not specified
if (-not $OutputCSVPath) {
    $OutputCSVPath = ".\PurchaseTracking_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

# Export to CSV
Write-Output "`n💾 Exporting to CSV..."
try {
    $csvData | Export-Csv -Path $OutputCSVPath -NoTypeInformation -Encoding UTF8
    Write-Output "   ✅ Successfully created: $OutputCSVPath"
    Write-Output "   📊 Rows exported: $($csvData.Count)"
} catch {
    Write-Error "Failed to export CSV: $_"
    exit 1
}

# Display summary
Write-Output "`n╔════════════════════════════════════════════════════════╗"
Write-Output "║                     SUMMARY                            ║"
Write-Output "╚════════════════════════════════════════════════════════╝"
Write-Output "Vendor:        $vendor"
Write-Output "Invoice:       $($parsedData.InvoiceNumber)"
Write-Output "Location:      $location"
Write-Output "Items:         $($csvData.Count)"
Write-Output "Output CSV:    $OutputCSVPath"
Write-Output "`n✅ Complete! You can now open the CSV in Excel or import to your system.`n"