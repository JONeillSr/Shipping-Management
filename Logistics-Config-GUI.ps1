<#
.SYNOPSIS
    GUI-based Logistics Email Configuration Tool with PDF Parsing & Template Manager
    
.DESCRIPTION
    Provides a graphical interface to configure auction logistics details,
    parse PDF invoices for automatic data extraction, manage reusable templates,
    and generate configuration files for the email automation script.
    
.PARAMETER PDFInvoice
    Optional path to PDF invoice for automatic data extraction
    
.PARAMETER LoadTemplate
    Load an existing template configuration
    
.EXAMPLE
    .\Logistics-Config-GUI.ps1
    
.EXAMPLE
    .\Logistics-Config-GUI.ps1 -PDFInvoice ".\invoice.pdf"
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 2025-01-07
    Version: 2.0.1
    Change Date: 2025-01-08
    Change Purpose: Fixed all PSScriptAnalyzer warnings (renamed functions: Load-Template→Import-Template, Refresh-TemplateList→Update-TemplateList, Load-TemplateToForm→Import-TemplateToForm)
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$PDFInvoice,
    
    [Parameter(Mandatory=$false)]
    [string]$LoadTemplate
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region PDF Parsing Functions
function Get-PDFText {
    <#
    .SYNOPSIS
        Extracts text from PDF using multiple fallback methods
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 2.0.1
        Change Date: 2025-01-08
        Change Purpose: Fixed PSScriptAnalyzer warnings - renamed from Extract-PDFText
    #>
    param (
        [string]$PDFPath
    )
    
    try {
        Write-Host "Extracting text from PDF: $PDFPath" -ForegroundColor Cyan
        
        # Method 1: Try using iTextSharp (if available)
        try {
            if (Get-Command "iTextSharp*" -ErrorAction SilentlyContinue) {
                Write-Host "  Using iTextSharp..." -ForegroundColor Gray
                # iTextSharp extraction would go here
            }
        }
        catch { }
        
        # Method 2: Try using System.IO.Packaging (works for some PDFs)
        try {
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $text = ""
            
            # Read PDF as binary
            $bytes = [System.IO.File]::ReadAllBytes($PDFPath)
            $pdfText = [System.Text.Encoding]::UTF8.GetString($bytes)
            
            # Extract readable text between stream markers
            $pattern = 'stream(.*?)endstream'
            $regexMatches = [regex]::Matches($pdfText, $pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
            
            foreach ($match in $regexMatches) {
                $streamContent = $match.Groups[1].Value
                # Try to extract readable text
                $readableText = [regex]::Replace($streamContent, '[^\x20-\x7E\r\n]', '')
                $text += $readableText + "`n"
            }
            
            if ($text.Length -gt 100) {
                Write-Host "  ✅ Extracted $($text.Length) characters" -ForegroundColor Green
                return $text
            }
        }
        catch {
            Write-Host "  Method 2 failed: $_" -ForegroundColor Yellow
        }
        
        # Method 3: Simple binary extraction for text-based PDFs
        try {
            Write-Host "  Trying binary text extraction..." -ForegroundColor Gray
            $pdfContent = [System.IO.File]::ReadAllText($PDFPath, [System.Text.Encoding]::GetEncoding('ISO-8859-1'))
            
            # Remove binary junk, keep readable text
            $cleanText = [regex]::Replace($pdfContent, '[^\x20-\x7E\r\n]+', ' ')
            $cleanText = [regex]::Replace($cleanText, '\s+', ' ')
            
            if ($cleanText.Length -gt 100) {
                Write-Host "  ✅ Extracted $($cleanText.Length) characters" -ForegroundColor Green
                return $cleanText
            }
        }
        catch {
            Write-Host "  Method 3 failed: $_" -ForegroundColor Yellow
        }
        
        # Method 4: Use external PDF to text tool if available
        $pdfToTextPath = "C:\Program Files\PDFtk\bin\pdftotext.exe"
        if (Test-Path $pdfToTextPath) {
            Write-Host "  Using PDFtoText utility..." -ForegroundColor Gray
            $tempTxtFile = [System.IO.Path]::GetTempFileName()
            & $pdfToTextPath $PDFPath $tempTxtFile
            if (Test-Path $tempTxtFile) {
                $text = Get-Content $tempTxtFile -Raw
                Remove-Item $tempTxtFile -Force
                return $text
            }
        }
        
        throw "All PDF extraction methods failed"
    }
    catch {
        Write-Warning "Could not extract PDF text: $_"
        Write-Host "`n💡 TIP: For best results, install a PDF text extraction tool or save invoice as text." -ForegroundColor Yellow
        return $null
    }
}

function Get-BrolynInvoiceData {
    <#
    .SYNOPSIS
        Specialized parser for Brolyn Auctions invoice format
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.1
        Change Date: 2025-01-08
        Change Purpose: Fixed PSScriptAnalyzer warnings - renamed from Parse-BrolynInvoice
    #>
    param (
        [string]$InvoiceText
    )
    
    $parsed = @{
        AuctionCompany = "Brolyn Auctions"
        ContactPhone = $null
        ContactEmail = $null
        PickupAddress = @()
        PickupDates = @()
        SpecialNotes = @()
    }
    
    if (!$InvoiceText) { return $parsed }
    
    # Extract Brolyn contact info
    if ($InvoiceText -match 'Phone:\s*\((\d{3})\)\s*(\d{3})-(\d{4})') {
        $parsed.ContactPhone = "($($Matches[1])) $($Matches[2])-$($Matches[3])"
    }
    elseif ($InvoiceText -match '(\d{3})\s*-\s*(\d{3})\s*-\s*(\d{4})') {
        $parsed.ContactPhone = "$($Matches[1])-$($Matches[2])-$($Matches[3])"
    }
    
    # Extract email
    if ($InvoiceText -match 'logistics@brolynauctions\.com') {
        $parsed.ContactEmail = "logistics@brolynauctions.com"
    }
    
    # Extract addresses - Brolyn typically has multiple locations
    if ($InvoiceText -match '290\s+West\s+750\s+North.*?Howe.*?IN\s+46746') {
        $parsed.PickupAddress += "290 West 750 North (Plant 208/209), Howe, IN 46746"
    }
    
    if ($InvoiceText -match '1139\s+Haines\s+Blvd.*?Sturgis.*?MI\s+49091') {
        $parsed.PickupAddress += "1139 Haines Blvd (Plant 901), Sturgis, MI 49091"
    }
    
    # Extract load times for materials
    if ($InvoiceText -match 'load times for materials.*?(\w+\s+\d{1,2}/\d{1,2})\s+thru\s+(\w+\s+\d{1,2}/\d{1,2}).*?(\d{1,2}[ap]m)\s*-\s*(\d{1,2}[ap]m)') {
        $startDate = $Matches[1]
        $endDate = $Matches[2]
        $startTime = $Matches[3]
        $endTime = $Matches[4]
        $parsed.PickupDates += "$startDate through $endDate, $startTime to $endTime EST"
        $parsed.SpecialNotes += "Materials (raw goods and RV components) must be picked up between specified dates"
    }
    
    # Extract load times for racking/equipment
    if ($InvoiceText -match 'load times for racking and equipment.*?(\w+\s+\d{1,2}/\d{1,2})\s+thru\s+(\w+\s+\d{1,2}/\d{1,2})') {
        $rackStart = $Matches[1]
        $rackEnd = $Matches[2]
        $parsed.SpecialNotes += "Racking and equipment pickup: $rackStart through $rackEnd"
    }
    
    # Extract special notes
    if ($InvoiceText -match 'responsibility of the buyer to arrange proper rigging') {
        $parsed.SpecialNotes += "Buyer responsible for arranging proper rigging for items requiring expert removal"
    }
    
    if ($InvoiceText -match 'Items not picked up.*?will be considered abandoned') {
        $parsed.SpecialNotes += "Items not picked up within specified windows will be considered abandoned"
    }
    
    return $parsed
}

function Get-InvoiceData {
    <#
    .SYNOPSIS
        Intelligently parses invoice text to extract logistics information
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 2.0.1
        Change Date: 2025-01-08
        Change Purpose: Fixed PSScriptAnalyzer warnings - renamed from Parse-InvoiceData
    #>
    param (
        [string]$InvoiceText
    )
    
    # Try vendor-specific parsers first
    if ($InvoiceText -match 'Brolyn') {
        Write-Host "  🎯 Detected Brolyn Auctions format" -ForegroundColor Green
        return Get-BrolynInvoiceData -InvoiceText $InvoiceText
    }
    
    # Generic parser for other vendors
    $parsedData = @{
        AuctionCompany = $null
        ContactPhone = $null
        ContactEmail = $null
        PickupAddress = $null
        PickupDates = @()
        SpecialNotes = @()
    }
    
    if (!$InvoiceText) { return $parsedData }
    
    # Extract phone numbers
    if ($InvoiceText -match '\((\d{3})\)\s*(\d{3})-(\d{4})') {
        $parsedData.ContactPhone = "($($Matches[1])) $($Matches[2])-$($Matches[3])"
    }
    elseif ($InvoiceText -match '(\d{3})[-\.\s](\d{3})[-\.\s](\d{4})') {
        $parsedData.ContactPhone = "$($Matches[1])-$($Matches[2])-$($Matches[3])"
    }
    
    # Extract email addresses
    if ($InvoiceText -match '([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})') {
        $parsedData.ContactEmail = $Matches[1]
    }
    
    # Extract addresses
    if ($InvoiceText -match '(\d+\s+[A-Za-z\s]+(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd|Drive|Dr|Lane|Ln|Way|Court|Ct|Highway|Hwy)[,\s]+[A-Za-z\s]+[,\s]+[A-Z]{2}\s+\d{5})') {
        $parsedData.PickupAddress = $Matches[1]
    }
    
    # Extract auction company name
    if ($InvoiceText -match '([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\s+Auctions?)') {
        $parsedData.AuctionCompany = $Matches[1]
    }
    
    # Extract dates
    $datePattern = '(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)[,\s]+\w+\s+\d{1,2}(?:st|nd|rd|th)?[,\s]+\d{4}'
    $allDateMatches = [regex]::Matches($InvoiceText, $datePattern)
    foreach ($dateMatch in $allDateMatches) {
        if ($parsedData.PickupDates -notcontains $dateMatch.Value) {
            $parsedData.PickupDates += $dateMatch.Value
        }
    }
    
    return $parsedData
}
#endregion

#region Template Management
function Get-TemplateList {
    <#
    .SYNOPSIS
        Gets list of available templates
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$TemplateDir = ".\Templates"
    )
    
    if (!(Test-Path $TemplateDir)) {
        New-Item -ItemType Directory -Path $TemplateDir -Force | Out-Null
    }
    
    $templates = Get-ChildItem -Path $TemplateDir -Filter "*.json" -ErrorAction SilentlyContinue
    return $templates
}

function Import-Template {
    <#
    .SYNOPSIS
        Loads template and populates form fields
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.1
        Change Date: 2025-01-08
        Change Purpose: Fixed PSScriptAnalyzer warning - renamed from Load-Template
    #>
    param (
        [string]$TemplatePath
    )
    
    try {
        $config = Get-Content $TemplatePath -Raw | ConvertFrom-Json
        Write-Host "✅ Loaded template: $(Split-Path $TemplatePath -Leaf)" -ForegroundColor Green
        return $config
    }
    catch {
        Write-Warning "Failed to load template: $_"
        return $null
    }
}

function Save-TemplateFile {
    <#
    .SYNOPSIS
        Saves configuration as template
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [object]$Config,
        [string]$TemplateName,
        [string]$TemplateDir = ".\Templates"
    )
    
    if (!(Test-Path $TemplateDir)) {
        New-Item -ItemType Directory -Path $TemplateDir -Force | Out-Null
    }
    
    $templatePath = Join-Path $TemplateDir "$TemplateName.json"
    $Config | ConvertTo-Json -Depth 10 | Out-File $templatePath -Encoding UTF8
    
    Write-Host "✅ Template saved: $templatePath" -ForegroundColor Green
    return $templatePath
}

function Remove-TemplateFile {
    <#
    .SYNOPSIS
        Deletes a template file
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$TemplatePath
    )
    
    if (Test-Path $TemplatePath) {
        Remove-Item $TemplatePath -Force
        Write-Host "🗑️ Template deleted: $(Split-Path $TemplatePath -Leaf)" -ForegroundColor Yellow
        return $true
    }
    return $false
}
#endregion

#region Configuration Presets
$script:CommonValues = @{
    AuctionCompanies = @(
        "Brolyn Auctions",
        "Ritchie Bros",
        "Purple Wave",
        "IronPlanet",
        "GovDeals",
        "Public Surplus",
        "Auctions International",
        "Custom Entry..."
    )
    
    TruckTypes = @(
        "53' dry van",
        "48' dry van",
        "Flatbed (tarped)",
        "Step deck",
        "Hotshot",
        "Box truck",
        "Two trucks: 53' dry van + Flatbed (tarped)",
        "Custom Entry..."
    )
    
    SpecialNotesPresets = @(
        "Forklift available on site for loading",
        "Loading dock available at pickup",
        "Freight prep and consolidation needed",
        "Loading assistance required (1-2 people)",
        "Delivery location has dock with pallet jacks only",
        "Trucks must back up to dock for unloading",
        "Driver must call 1 hour before delivery",
        "Total weight will NOT exceed standard truck capacity",
        "Rigging may be required for heavy items",
        "Items must be picked up within specified window or considered abandoned",
        "Custom Entry..."
    )
    
    DeliveryAddresses = @(
        "1218 Lake Avenue, Ashtabula, OH 44004",
        "Custom Entry..."
    )
}
#endregion

#region GUI Creation
function New-ConfigurationGUI {
    <#
    .SYNOPSIS
        Creates the main configuration GUI with Template Manager
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 2.0.1
        Change Date: 2025-01-08
        Change Purpose: Fixed PSScriptAnalyzer warnings - updated helper function names
    #>
    
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Logistics Email Configuration Tool v2.0.1"
    $form.Size = New-Object System.Drawing.Size(1100, 800)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    #region Template Manager Panel (Left Side)
    $pnlTemplateManager = New-Object System.Windows.Forms.Panel
    $pnlTemplateManager.Location = New-Object System.Drawing.Point(10, 10)
    $pnlTemplateManager.Size = New-Object System.Drawing.Size(220, 680)
    $pnlTemplateManager.BorderStyle = "FixedSingle"
    $pnlTemplateManager.BackColor = [System.Drawing.Color]::WhiteSmoke
    
    # Template Manager Header
    $lblTemplateHeader = New-Object System.Windows.Forms.Label
    $lblTemplateHeader.Location = New-Object System.Drawing.Point(10, 10)
    $lblTemplateHeader.Size = New-Object System.Drawing.Size(200, 25)
    $lblTemplateHeader.Text = "📚 Template Manager"
    $lblTemplateHeader.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $lblTemplateHeader.ForeColor = [System.Drawing.Color]::DarkBlue
    $pnlTemplateManager.Controls.Add($lblTemplateHeader)
    
    # Template ListBox
    $lstTemplates = New-Object System.Windows.Forms.ListBox
    $lstTemplates.Location = New-Object System.Drawing.Point(10, 45)
    $lstTemplates.Size = New-Object System.Drawing.Size(200, 500)
    $pnlTemplateManager.Controls.Add($lstTemplates)
    
    # Template Buttons
    $btnLoadTemplate = New-Object System.Windows.Forms.Button
    $btnLoadTemplate.Location = New-Object System.Drawing.Point(10, 555)
    $btnLoadTemplate.Size = New-Object System.Drawing.Size(200, 30)
    $btnLoadTemplate.Text = "📂 Load Selected"
    $btnLoadTemplate.BackColor = [System.Drawing.Color]::LightGreen
    $pnlTemplateManager.Controls.Add($btnLoadTemplate)
    
    $btnDeleteTemplate = New-Object System.Windows.Forms.Button
    $btnDeleteTemplate.Location = New-Object System.Drawing.Point(10, 590)
    $btnDeleteTemplate.Size = New-Object System.Drawing.Size(200, 30)
    $btnDeleteTemplate.Text = "🗑️ Delete Selected"
    $btnDeleteTemplate.BackColor = [System.Drawing.Color]::LightCoral
    $pnlTemplateManager.Controls.Add($btnDeleteTemplate)
    
    $btnRefreshTemplates = New-Object System.Windows.Forms.Button
    $btnRefreshTemplates.Location = New-Object System.Drawing.Point(10, 625)
    $btnRefreshTemplates.Size = New-Object System.Drawing.Size(200, 30)
    $btnRefreshTemplates.Text = "🔄 Refresh List"
    $pnlTemplateManager.Controls.Add($btnRefreshTemplates)
    
    $form.Controls.Add($pnlTemplateManager)
    #endregion
    
    # Create TabControl (Right Side)
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(240, 10)
    $tabControl.Size = New-Object System.Drawing.Size(835, 680)
    
    #region Tab 1: Basic Info
    $tabBasic = New-Object System.Windows.Forms.TabPage
    $tabBasic.Text = "Basic Information"
    $tabBasic.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $yPos = 20
    
    # PDF Import Button
    $btnImportPDF = New-Object System.Windows.Forms.Button
    $btnImportPDF.Location = New-Object System.Drawing.Point(20, $yPos)
    $btnImportPDF.Size = New-Object System.Drawing.Size(200, 35)
    $btnImportPDF.Text = "📄 Import from PDF Invoice"
    $btnImportPDF.BackColor = [System.Drawing.Color]::LightBlue
    $btnImportPDF.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $tabBasic.Controls.Add($btnImportPDF)
    
    $yPos += 50
    
    # Auction Company
    $lblAuction = New-Object System.Windows.Forms.Label
    $lblAuction.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblAuction.Size = New-Object System.Drawing.Size(200, 20)
    $lblAuction.Text = "Auction Company:"
    $tabBasic.Controls.Add($lblAuction)
    
    $cmbAuction = New-Object System.Windows.Forms.ComboBox
    $cmbAuction.Location = New-Object System.Drawing.Point(230, $yPos)
    $cmbAuction.Size = New-Object System.Drawing.Size(400, 25)
    $cmbAuction.DropDownStyle = "DropDown"
    $script:CommonValues.AuctionCompanies | ForEach-Object { $cmbAuction.Items.Add($_) | Out-Null }
    $tabBasic.Controls.Add($cmbAuction)
    
    $yPos += 40
    
    # Pickup Address
    $lblPickup = New-Object System.Windows.Forms.Label
    $lblPickup.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblPickup.Size = New-Object System.Drawing.Size(200, 20)
    $lblPickup.Text = "Pickup Address:"
    $tabBasic.Controls.Add($lblPickup)
    
    $txtPickup = New-Object System.Windows.Forms.TextBox
    $txtPickup.Location = New-Object System.Drawing.Point(230, $yPos)
    $txtPickup.Size = New-Object System.Drawing.Size(500, 25)
    $txtPickup.Multiline = $true
    $txtPickup.Height = 50
    $txtPickup.ScrollBars = "Vertical"
    $tabBasic.Controls.Add($txtPickup)
    
    $yPos += 70
    
    # Delivery Address
    $lblDelivery = New-Object System.Windows.Forms.Label
    $lblDelivery.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblDelivery.Size = New-Object System.Drawing.Size(200, 20)
    $lblDelivery.Text = "Delivery Address:"
    $tabBasic.Controls.Add($lblDelivery)
    
    $cmbDelivery = New-Object System.Windows.Forms.ComboBox
    $cmbDelivery.Location = New-Object System.Drawing.Point(230, $yPos)
    $cmbDelivery.Size = New-Object System.Drawing.Size(500, 25)
    $cmbDelivery.DropDownStyle = "DropDown"
    $script:CommonValues.DeliveryAddresses | ForEach-Object { $cmbDelivery.Items.Add($_) | Out-Null }
    $cmbDelivery.Text = "1218 Lake Avenue, Ashtabula, OH 44004"
    $tabBasic.Controls.Add($cmbDelivery)
    
    $yPos += 40
    
    # Contact Phone
    $lblPhone = New-Object System.Windows.Forms.Label
    $lblPhone.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblPhone.Size = New-Object System.Drawing.Size(200, 20)
    $lblPhone.Text = "Logistics Contact Phone:"
    $tabBasic.Controls.Add($lblPhone)
    
    $txtPhone = New-Object System.Windows.Forms.TextBox
    $txtPhone.Location = New-Object System.Drawing.Point(230, $yPos)
    $txtPhone.Size = New-Object System.Drawing.Size(200, 25)
    $tabBasic.Controls.Add($txtPhone)
    
    $yPos += 40
    
    # Contact Email
    $lblEmail = New-Object System.Windows.Forms.Label
    $lblEmail.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblEmail.Size = New-Object System.Drawing.Size(200, 20)
    $lblEmail.Text = "Logistics Contact Email:"
    $tabBasic.Controls.Add($lblEmail)
    
    $txtEmail = New-Object System.Windows.Forms.TextBox
    $txtEmail.Location = New-Object System.Drawing.Point(230, $yPos)
    $txtEmail.Size = New-Object System.Drawing.Size(300, 25)
    $tabBasic.Controls.Add($txtEmail)
    
    $yPos += 40
    
    # Pickup Date/Time
    $lblPickupDate = New-Object System.Windows.Forms.Label
    $lblPickupDate.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblPickupDate.Size = New-Object System.Drawing.Size(200, 20)
    $lblPickupDate.Text = "Pickup Date/Time:"
    $tabBasic.Controls.Add($lblPickupDate)
    
    $txtPickupDate = New-Object System.Windows.Forms.TextBox
    $txtPickupDate.Location = New-Object System.Drawing.Point(230, $yPos)
    $txtPickupDate.Size = New-Object System.Drawing.Size(500, 25)
    $txtPickupDate.Text = "Thursday October 9th, 2025 12:00pm EST"
    $tabBasic.Controls.Add($txtPickupDate)
    
    $yPos += 40
    
    # Delivery Date/Time
    $lblDeliveryDate = New-Object System.Windows.Forms.Label
    $lblDeliveryDate.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblDeliveryDate.Size = New-Object System.Drawing.Size(200, 20)
    $lblDeliveryDate.Text = "Delivery Date/Time:"
    $tabBasic.Controls.Add($lblDeliveryDate)
    
    $txtDeliveryDate = New-Object System.Windows.Forms.TextBox
    $txtDeliveryDate.Location = New-Object System.Drawing.Point(230, $yPos)
    $txtDeliveryDate.Size = New-Object System.Drawing.Size(500, 25)
    $txtDeliveryDate.Text = "Friday October 10th, 2025 between 9:00am and 5:00pm EST"
    $tabBasic.Controls.Add($txtDeliveryDate)
    
    $yPos += 40
    
    # Delivery Notice
    $lblNotice = New-Object System.Windows.Forms.Label
    $lblNotice.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblNotice.Size = New-Object System.Drawing.Size(200, 20)
    $lblNotice.Text = "Delivery Notice Requirements:"
    $tabBasic.Controls.Add($lblNotice)
    
    $txtNotice = New-Object System.Windows.Forms.TextBox
    $txtNotice.Location = New-Object System.Drawing.Point(230, $yPos)
    $txtNotice.Size = New-Object System.Drawing.Size(500, 25)
    $txtNotice.Text = "Driver must call at least one hour prior to delivery"
    $tabBasic.Controls.Add($txtNotice)
    
    $tabControl.TabPages.Add($tabBasic)
    #endregion
    
    #region Tab 2: Shipping Details
    $tabShipping = New-Object System.Windows.Forms.TabPage
    $tabShipping.Text = "Shipping Details"
    $tabShipping.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $yPos = 20
    
    # Total Pallets
    $lblPallets = New-Object System.Windows.Forms.Label
    $lblPallets.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblPallets.Size = New-Object System.Drawing.Size(200, 20)
    $lblPallets.Text = "Total Pallets:"
    $tabShipping.Controls.Add($lblPallets)
    
    $txtPallets = New-Object System.Windows.Forms.TextBox
    $txtPallets.Location = New-Object System.Drawing.Point(230, $yPos)
    $txtPallets.Size = New-Object System.Drawing.Size(100, 25)
    $txtPallets.Text = "TBD"
    $tabShipping.Controls.Add($txtPallets)
    
    $yPos += 40
    
    # Truck Types
    $lblTrucks = New-Object System.Windows.Forms.Label
    $lblTrucks.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblTrucks.Size = New-Object System.Drawing.Size(200, 20)
    $lblTrucks.Text = "Truck Types Needed:"
    $tabShipping.Controls.Add($lblTrucks)
    
    $cmbTrucks = New-Object System.Windows.Forms.ComboBox
    $cmbTrucks.Location = New-Object System.Drawing.Point(230, $yPos)
    $cmbTrucks.Size = New-Object System.Drawing.Size(550, 25)
    $cmbTrucks.DropDownStyle = "DropDown"
    $script:CommonValues.TruckTypes | ForEach-Object { $cmbTrucks.Items.Add($_) | Out-Null }
    $cmbTrucks.Text = "Two trucks: 53' dry van + Flatbed (tarped)"
    $tabShipping.Controls.Add($cmbTrucks)
    
    $yPos += 40
    
    # Labor Requirements
    $lblLabor = New-Object System.Windows.Forms.Label
    $lblLabor.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblLabor.Size = New-Object System.Drawing.Size(200, 20)
    $lblLabor.Text = "Labor Requirements:"
    $tabShipping.Controls.Add($lblLabor)
    
    $txtLabor = New-Object System.Windows.Forms.TextBox
    $txtLabor.Location = New-Object System.Drawing.Point(230, $yPos)
    $txtLabor.Size = New-Object System.Drawing.Size(550, 25)
    $txtLabor.Text = "1-2 people for consolidation, freight prep, and loading"
    $tabShipping.Controls.Add($txtLabor)
    
    $yPos += 40
    
    # Weight Notes
    $lblWeight = New-Object System.Windows.Forms.Label
    $lblWeight.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblWeight.Size = New-Object System.Drawing.Size(200, 20)
    $lblWeight.Text = "Weight Notes:"
    $tabShipping.Controls.Add($lblWeight)
    
    $txtWeight = New-Object System.Windows.Forms.TextBox
    $txtWeight.Location = New-Object System.Drawing.Point(230, $yPos)
    $txtWeight.Size = New-Object System.Drawing.Size(550, 25)
    $txtWeight.Text = "Total weight unknown but will NOT exceed standard load capacity"
    $tabShipping.Controls.Add($txtWeight)
    
    $yPos += 60
    
    # Special Notes Section
    $lblSpecialNotes = New-Object System.Windows.Forms.Label
    $lblSpecialNotes.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblSpecialNotes.Size = New-Object System.Drawing.Size(200, 20)
    $lblSpecialNotes.Text = "Special Notes:"
    $lblSpecialNotes.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $tabShipping.Controls.Add($lblSpecialNotes)
    
    $yPos += 30
    
    # Special Notes Preset Selector
    $cmbSpecialNotes = New-Object System.Windows.Forms.ComboBox
    $cmbSpecialNotes.Location = New-Object System.Drawing.Point(20, $yPos)
    $cmbSpecialNotes.Size = New-Object System.Drawing.Size(650, 25)
    $cmbSpecialNotes.DropDownStyle = "DropDownList"
    $script:CommonValues.SpecialNotesPresets | ForEach-Object { $cmbSpecialNotes.Items.Add($_) | Out-Null }
    $tabShipping.Controls.Add($cmbSpecialNotes)
    
    $btnAddNote = New-Object System.Windows.Forms.Button
    $btnAddNote.Location = New-Object System.Drawing.Point(680, $yPos)
    $btnAddNote.Size = New-Object System.Drawing.Size(100, 25)
    $btnAddNote.Text = "Add Note"
    $tabShipping.Controls.Add($btnAddNote)
    
    $yPos += 35
    
    # Special Notes ListBox
    $lstSpecialNotes = New-Object System.Windows.Forms.ListBox
    $lstSpecialNotes.Location = New-Object System.Drawing.Point(20, $yPos)
    $lstSpecialNotes.Size = New-Object System.Drawing.Size(760, 200)
    $tabShipping.Controls.Add($lstSpecialNotes)
    
    $btnRemoveNote = New-Object System.Windows.Forms.Button
    $btnRemoveNote.Location = New-Object System.Drawing.Point(20, ($yPos + 210))
    $btnRemoveNote.Size = New-Object System.Drawing.Size(150, 25)
    $btnRemoveNote.Text = "Remove Selected"
    $tabShipping.Controls.Add($btnRemoveNote)
    
    $btnClearNotes = New-Object System.Windows.Forms.Button
    $btnClearNotes.Location = New-Object System.Drawing.Point(180, ($yPos + 210))
    $btnClearNotes.Size = New-Object System.Drawing.Size(100, 25)
    $btnClearNotes.Text = "Clear All"
    $tabShipping.Controls.Add($btnClearNotes)
    
    $tabControl.TabPages.Add($tabShipping)
    #endregion
    
    #region Tab 3: Subject Line & Preview
    $tabPreview = New-Object System.Windows.Forms.TabPage
    $tabPreview.Text = "Subject & Preview"
    $tabPreview.Padding = New-Object System.Windows.Forms.Padding(10)
    
    $yPos = 20
    
    # Subject Line Generator
    $grpSubject = New-Object System.Windows.Forms.GroupBox
    $grpSubject.Location = New-Object System.Drawing.Point(20, $yPos)
    $grpSubject.Size = New-Object System.Drawing.Size(780, 120)
    $grpSubject.Text = "Email Subject Line Generator"
    
    $lblSubjectTemplate = New-Object System.Windows.Forms.Label
    $lblSubjectTemplate.Location = New-Object System.Drawing.Point(15, 30)
    $lblSubjectTemplate.Size = New-Object System.Drawing.Size(750, 20)
    $lblSubjectTemplate.Text = "Format: Freight Quote Request - [Pickup City, ST] to [Delivery City, ST] - Pickup [Date]"
    $lblSubjectTemplate.ForeColor = [System.Drawing.Color]::Gray
    $grpSubject.Controls.Add($lblSubjectTemplate)
    
    $txtSubject = New-Object System.Windows.Forms.TextBox
    $txtSubject.Location = New-Object System.Drawing.Point(15, 55)
    $txtSubject.Size = New-Object System.Drawing.Size(750, 25)
    $txtSubject.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $grpSubject.Controls.Add($txtSubject)
    
    $btnGenerateSubject = New-Object System.Windows.Forms.Button
    $btnGenerateSubject.Location = New-Object System.Drawing.Point(15, 85)
    $btnGenerateSubject.Size = New-Object System.Drawing.Size(180, 25)
    $btnGenerateSubject.Text = "🔄 Auto-Generate Subject"
    $btnGenerateSubject.BackColor = [System.Drawing.Color]::LightGreen
    $grpSubject.Controls.Add($btnGenerateSubject)
    
    $tabPreview.Controls.Add($grpSubject)
    
    $yPos += 140
    
    # Preview Section
    $lblPreview = New-Object System.Windows.Forms.Label
    $lblPreview.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblPreview.Size = New-Object System.Drawing.Size(200, 20)
    $lblPreview.Text = "Configuration Preview:"
    $lblPreview.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $tabPreview.Controls.Add($lblPreview)
    
    $yPos += 30
    
    $txtPreview = New-Object System.Windows.Forms.TextBox
    $txtPreview.Location = New-Object System.Drawing.Point(20, $yPos)
    $txtPreview.Size = New-Object System.Drawing.Size(780, 420)
    $txtPreview.Multiline = $true
    $txtPreview.ScrollBars = "Vertical"
    $txtPreview.ReadOnly = $true
    $txtPreview.Font = New-Object System.Drawing.Font("Consolas", 9)
    $tabPreview.Controls.Add($txtPreview)
    
    $tabControl.TabPages.Add($tabPreview)
    #endregion
    
    $form.Controls.Add($tabControl)
    
    #region Bottom Buttons
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(750, 700)
    $btnSave.Size = New-Object System.Drawing.Size(150, 35)
    $btnSave.Text = "💾 Save Configuration"
    $btnSave.BackColor = [System.Drawing.Color]::LightGreen
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnSave)
    
    $btnSaveTemplate = New-Object System.Windows.Forms.Button
    $btnSaveTemplate.Location = New-Object System.Drawing.Point(910, 700)
    $btnSaveTemplate.Size = New-Object System.Drawing.Size(160, 35)
    $btnSaveTemplate.Text = "📋 Save as Template"
    $btnSaveTemplate.BackColor = [System.Drawing.Color]::LightBlue
    $btnSaveTemplate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnSaveTemplate)
    #endregion
    
    #region Helper Functions
    function Update-TemplateList {
        $lstTemplates.Items.Clear()
        $templates = Get-TemplateList
        foreach ($template in $templates) {
            $lstTemplates.Items.Add($template.BaseName) | Out-Null
        }
        
        if ($templates.Count -eq 0) {
            $lstTemplates.Items.Add("(No templates found)") | Out-Null
        }
    }
    
    function Import-TemplateToForm {
        param([object]$Config)
        
        if ($Config.auction_info.auction_name) { $cmbAuction.Text = $Config.auction_info.auction_name }
        if ($Config.auction_info.pickup_address) { $txtPickup.Text = $Config.auction_info.pickup_address }
        if ($Config.delivery_address) { $cmbDelivery.Text = $Config.delivery_address }
        if ($Config.auction_info.logistics_contact.phone) { $txtPhone.Text = $Config.auction_info.logistics_contact.phone }
        if ($Config.auction_info.logistics_contact.email) { $txtEmail.Text = $Config.auction_info.logistics_contact.email }
        if ($Config.auction_info.pickup_datetime) { $txtPickupDate.Text = $Config.auction_info.pickup_datetime }
        if ($Config.auction_info.delivery_datetime) { $txtDeliveryDate.Text = $Config.auction_info.delivery_datetime }
        if ($Config.auction_info.delivery_notice) { $txtNotice.Text = $Config.auction_info.delivery_notice }
        
        if ($Config.shipping_requirements.total_pallets) { $txtPallets.Text = $Config.shipping_requirements.total_pallets }
        if ($Config.shipping_requirements.truck_types) { $cmbTrucks.Text = $Config.shipping_requirements.truck_types }
        if ($Config.shipping_requirements.labor_needed) { $txtLabor.Text = $Config.shipping_requirements.labor_needed }
        if ($Config.shipping_requirements.weight_notes) { $txtWeight.Text = $Config.shipping_requirements.weight_notes }
        
        $lstSpecialNotes.Items.Clear()
        if ($Config.auction_info.special_notes) {
            foreach ($note in $Config.auction_info.special_notes) {
                $lstSpecialNotes.Items.Add($note) | Out-Null
            }
        }
        
        if ($Config.email_subject) { $txtSubject.Text = $Config.email_subject }
    }
    
    function Build-ConfigObject {
        $specialNotes = @()
        foreach ($item in $lstSpecialNotes.Items) {
            $specialNotes += $item.ToString()
        }
        
        $config = @{
            email_subject = $txtSubject.Text
            auction_info = @{
                auction_name = $cmbAuction.Text
                pickup_address = $txtPickup.Text
                logistics_contact = @{
                    phone = $txtPhone.Text
                    email = $txtEmail.Text
                }
                pickup_datetime = $txtPickupDate.Text
                delivery_datetime = $txtDeliveryDate.Text
                delivery_notice = $txtNotice.Text
                special_notes = $specialNotes
            }
            delivery_address = $cmbDelivery.Text
            shipping_requirements = @{
                total_pallets = $txtPallets.Text
                truck_types = $cmbTrucks.Text
                labor_needed = $txtLabor.Text
                weight_notes = $txtWeight.Text
            }
        }
        
        $txtPreview.Text = $config | ConvertTo-Json -Depth 10
        return $config
    }
    #endregion
    
    #region Event Handlers
    
    # Initial template list load
    Update-TemplateList
    
    # Refresh Templates Button
    $btnRefreshTemplates.Add_Click({
        Update-TemplateList
        [System.Windows.Forms.MessageBox]::Show(
            "Template list refreshed!",
            "Refresh Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    
    # Load Template Button
    $btnLoadTemplate.Add_Click({
        if ($lstTemplates.SelectedItem -and $lstTemplates.SelectedItem -ne "(No templates found)") {
            $templateName = $lstTemplates.SelectedItem
            $templatePath = Join-Path ".\Templates" "$templateName.json"
            $config = Import-Template -TemplatePath $templatePath
            
            if ($config) {
                Import-TemplateToForm -Config $config
                [System.Windows.Forms.MessageBox]::Show(
                    "Template loaded successfully!`n`nTemplate: $templateName",
                    "Template Loaded",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Please select a template to load.",
                "No Template Selected",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
    })
    
    # Delete Template Button
    $btnDeleteTemplate.Add_Click({
        if ($lstTemplates.SelectedItem -and $lstTemplates.SelectedItem -ne "(No templates found)") {
            $templateName = $lstTemplates.SelectedItem
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Are you sure you want to delete template:`n`n$templateName",
                "Confirm Delete",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            
            if ($result -eq "Yes") {
                $templatePath = Join-Path ".\Templates" "$templateName.json"
                Remove-TemplateFile -TemplatePath $templatePath
                Update-TemplateList
                
                [System.Windows.Forms.MessageBox]::Show(
                    "Template deleted successfully!",
                    "Delete Complete",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
    })
    
    # Import PDF Button
    $btnImportPDF.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
        $openFileDialog.Title = "Select Auction Invoice PDF"
        
        if ($openFileDialog.ShowDialog() -eq "OK") {
            Write-Host "`n📄 Processing PDF: $($openFileDialog.FileName)" -ForegroundColor Cyan
            
            $pdfText = Get-PDFText -PDFPath $openFileDialog.FileName
            
            if ($pdfText) {
                $parsed = Get-InvoiceData -InvoiceText $pdfText
                
                Write-Host "`nExtracted Data:" -ForegroundColor Yellow
                Write-Host "  Company: $($parsed.AuctionCompany)" -ForegroundColor White
                Write-Host "  Phone: $($parsed.ContactPhone)" -ForegroundColor White
                Write-Host "  Email: $($parsed.ContactEmail)" -ForegroundColor White
                Write-Host "  Addresses: $($parsed.PickupAddress.Count)" -ForegroundColor White
                Write-Host "  Dates: $($parsed.PickupDates.Count)" -ForegroundColor White
                Write-Host "  Notes: $($parsed.SpecialNotes.Count)" -ForegroundColor White
                
                # Populate form
                if ($parsed.AuctionCompany) { $cmbAuction.Text = $parsed.AuctionCompany }
                if ($parsed.ContactPhone) { $txtPhone.Text = $parsed.ContactPhone }
                if ($parsed.ContactEmail) { $txtEmail.Text = $parsed.ContactEmail }
                
                if ($parsed.PickupAddress -is [array] -and $parsed.PickupAddress.Count -gt 0) {
                    $txtPickup.Text = $parsed.PickupAddress -join "`r`n"
                }
                elseif ($parsed.PickupAddress) {
                    $txtPickup.Text = $parsed.PickupAddress
                }
                
                if ($parsed.PickupDates -and $parsed.PickupDates.Count -gt 0) {
                    $txtPickupDate.Text = $parsed.PickupDates[0]
                }
                
                $lstSpecialNotes.Items.Clear()
                foreach ($note in $parsed.SpecialNotes) {
                    $lstSpecialNotes.Items.Add($note) | Out-Null
                }
                
                [System.Windows.Forms.MessageBox]::Show(
                    "PDF data imported successfully!`n`nPlease review and adjust as needed.`n`nExtracted:`n- Company: $($parsed.AuctionCompany)`n- Phone: $($parsed.ContactPhone)`n- Email: $($parsed.ContactEmail)`n- Addresses: $($parsed.PickupAddress.Count)`n- Special Notes: $($parsed.SpecialNotes.Count)",
                    "Import Complete",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Could not extract text from PDF.`n`nPlease enter information manually.`n`nTip: Try saving the invoice as a text-based PDF if possible.",
                    "Import Failed",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        }
    })
    
    # Add Special Note Button
    $btnAddNote.Add_Click({
        $selectedNote = $cmbSpecialNotes.SelectedItem
        if ($selectedNote -and $selectedNote -ne "Custom Entry...") {
            if ($lstSpecialNotes.Items.Contains($selectedNote) -eq $false) {
                $lstSpecialNotes.Items.Add($selectedNote)
            }
        }
        elseif ($selectedNote -eq "Custom Entry...") {
            $customNote = [Microsoft.VisualBasic.Interaction]::InputBox(
                "Enter custom special note:",
                "Custom Note Entry",
                ""
            )
            if ($customNote) {
                $lstSpecialNotes.Items.Add($customNote)
            }
        }
    })
    
    # Remove Special Note Button
    $btnRemoveNote.Add_Click({
        if ($lstSpecialNotes.SelectedIndex -ge 0) {
            $lstSpecialNotes.Items.RemoveAt($lstSpecialNotes.SelectedIndex)
        }
    })
    
    # Clear Notes Button
    $btnClearNotes.Add_Click({
        $lstSpecialNotes.Items.Clear()
    })
    
    # Generate Subject Line Button
    $btnGenerateSubject.Add_Click({
        try {
            # Parse pickup address for city/state
            $pickupText = $txtPickup.Text
            $pickupCity = "Unknown"
            $pickupState = "XX"
            
            # Try multiple patterns for city/state extraction
            if ($pickupText -match ',\s*([A-Za-z\s]+),\s*([A-Z]{2})\s+\d{5}') {
                $pickupCity = $Matches[1].Trim()
                $pickupState = $Matches[2].Trim()
            }
            elseif ($pickupText -match '\(([A-Za-z\s]+)\).*?([A-Z]{2})') {
                $pickupCity = $Matches[1].Trim()
                $pickupState = $Matches[2].Trim()
            }
            
            # Parse delivery address for city/state
            $deliveryText = $cmbDelivery.Text
            $deliveryCity = "Unknown"
            $deliveryState = "XX"
            
            if ($deliveryText -match ',\s*([A-Za-z\s]+),\s*([A-Z]{2})') {
                $deliveryCity = $Matches[1].Trim()
                $deliveryState = $Matches[2].Trim()
            }
            
            # Parse pickup date
            $pickupDateText = $txtPickupDate.Text
            $pickupDate = "TBD"
            
            if ($pickupDateText -match '(\d{1,2}/\d{1,2}/\d{4})') {
                $pickupDate = $Matches[1]
            }
            elseif ($pickupDateText -match '([A-Za-z]+\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})') {
                try {
                    $dateStr = $Matches[1] -replace '(?:st|nd|rd|th)', ''
                    $dateObj = [DateTime]::Parse($dateStr)
                    $pickupDate = $dateObj.ToString("MM/dd/yyyy")
                }
                catch {
                    $pickupDate = $Matches[1]
                }
            }
            
            # Generate subject line
            $subject = "Freight Quote Request - $pickupCity, $pickupState to $deliveryCity, $deliveryState - Pickup $pickupDate"
            $txtSubject.Text = $subject
            
            [System.Windows.Forms.MessageBox]::Show(
                "Subject line generated!`n`nSubject:`n$subject`n`nPlease review and edit if needed.",
                "Subject Generated",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            Write-Warning "Subject generation error: $_"
            [System.Windows.Forms.MessageBox]::Show(
                "Could not auto-generate subject line.`n`nError: $_`n`nPlease enter manually.",
                "Generation Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
    })
    
    # Save Configuration Button
    $btnSave.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "JSON files (*.json)|*.json"
        $saveFileDialog.Title = "Save Configuration"
        $saveFileDialog.FileName = "auction_config.json"
        
        if ($saveFileDialog.ShowDialog() -eq "OK") {
            $config = Build-ConfigObject
            $config | ConvertTo-Json -Depth 10 | Out-File $saveFileDialog.FileName -Encoding UTF8
            
            [System.Windows.Forms.MessageBox]::Show(
                "Configuration saved successfully!`n`nFile: $($saveFileDialog.FileName)",
                "Save Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })
    
    # Save Template Button
    $btnSaveTemplate.Add_Click({
        $templateName = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter template name:",
            "Save Template",
            $cmbAuction.Text
        )
        
        if ($templateName) {
            $config = Build-ConfigObject
            $templatePath = Save-TemplateFile -Config $config -TemplateName $templateName
            
            Update-TemplateList
            
            [System.Windows.Forms.MessageBox]::Show(
                "Template saved successfully!`n`nTemplate: $templateName`nLocation: $templatePath`n`nYou can now quickly load this configuration from the Template Manager.",
                "Template Saved",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })
    
    # Update preview when switching to preview tab
    $tabControl.Add_SelectedIndexChanged({
        if ($tabControl.SelectedTab -eq $tabPreview) {
            Build-ConfigObject | Out-Null
        }
    })
    
    #endregion
    
    # Show form
    $form.Add_Shown({$form.Activate()})
    [void]$form.ShowDialog()
}
#endregion

#region Main Execution
# Load Visual Basic for InputBox
Add-Type -AssemblyName Microsoft.VisualBasic

# Launch GUI
Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   LOGISTICS EMAIL CONFIGURATION TOOL v2.0.1           ║" -ForegroundColor Cyan
Write-Host "║   with PDF Parsing & Template Manager                ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

# Auto-load PDF if provided
if ($PDFInvoice -and (Test-Path $PDFInvoice)) {
    Write-Host "Auto-loading PDF: $PDFInvoice" -ForegroundColor Yellow
    $script:AutoLoadPDF = $PDFInvoice
}

New-ConfigurationGUI
#endregion
