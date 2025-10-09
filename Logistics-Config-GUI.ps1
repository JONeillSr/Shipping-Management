<#
.SYNOPSIS
    GUI-based Logistics Email Configuration Tool with PDF Parsing & Template Manager
    
.DESCRIPTION
    Provides a graphical interface to configure auction logistics details,
    parse PDF invoices using Generic-PDF-Invoice-Parser.ps1 for automatic data extraction,
    manage reusable templates, and generate configuration files for email automation.
    
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
    Create Date: 01/07/2025
    Version: 2.3.2
    Change Date: 10/09/2025
    Change Purpose: Removed emojis from GUI buttons for universal compatibility

.CHANGELOG
    v2.3.2 (10/09/2025): Removed emojis from GUI buttons for universal compatibility
      - All buttons: Replaced emoji prefixes with clear text labels
      - Template Manager header: Removed emoji for consistent rendering
      - Ensures proper display across all Windows versions and fonts
      
    v2.3.1 (10/09/2025): Enhanced date parsing for subject line generation
      - Generate-Subject: Now handles date ranges (e.g., "Tuesday 10/7 thru Friday 10/10")
      - Generate-Subject: Added automatic year addition for dates without years
      - Generate-Subject: Supports 5 date format patterns including ranges with day names
      - Generate-Subject: Extracts first date from ranges as pickup date
      
    v2.3.0 (10/09/2025): Added ZIP codes to subject line (freight industry best practice)
      - Generate-Subject: Now extracts and includes ZIP codes for accurate routing
      - Generate-Subject: Enhanced to parse ZIP codes from multiline addresses
      - Updated subject format template to show ZIP code format
      - Added freight routing benefits to success message
      
    v2.2.0 (10/09/2025): Fixed structured address handling
      - NEW: Import-ParsedDataToForm function to handle structured address objects
      - NEW: Show-AddressSelectionDialog function for multiple pickup locations
      - Import-ParsedDataToForm: Formats multiline addresses with Address2 support
      - Enhanced import summary to show detailed address extraction
      
    v2.1.0 (10/09/2025): Integrated with Generic-PDF-Invoice-Parser.ps1
      - NEW: Invoke-PDFInvoiceParser function to call parser with -ReturnObject
      - Import PDF button now uses advanced parser instead of basic extraction
      - Proper structured data handling from parser output
      
    v2.0.1 (01/08/2025): Fixed PSScriptAnalyzer warnings
      - Renamed functions: Load-Template → Import-Template
      - Renamed functions: Refresh-TemplateList → Update-TemplateList
      - Renamed functions: Load-TemplateToForm → Import-TemplateToForm
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
Add-Type -AssemblyName Microsoft.VisualBasic

#region PDF Parsing Integration
function Invoke-PDFInvoiceParser {
    <#
    .SYNOPSIS
        Calls Generic-PDF-Invoice-Parser.ps1 and returns parsed data object
    
    .DESCRIPTION
        Integrates with the Generic-PDF-Invoice-Parser.ps1 script by calling it with
        the -ReturnObject parameter to retrieve structured invoice data that can be
        mapped to the GUI form fields.
    
    .PARAMETER PDFPath
        Full path to the PDF invoice file to parse
    
    .EXAMPLE
        $data = Invoke-PDFInvoiceParser -PDFPath ".\invoice.pdf"
    
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 10/09/2025
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$PDFPath
    )
    
    try {
        # Check if parser script exists
        $parserScript = Join-Path $PSScriptRoot "Generic-PDF-Invoice-Parser.ps1"
        
        if (-not (Test-Path $parserScript)) {
            Write-Warning "Generic-PDF-Invoice-Parser.ps1 not found in script directory"
            Write-Host "Looking for: $parserScript" -ForegroundColor Yellow
            return $null
        }
        
        Write-Host "`n📄 Calling Generic-PDF-Invoice-Parser.ps1..." -ForegroundColor Cyan
        Write-Host "   Parser: $parserScript" -ForegroundColor Gray
        Write-Host "   PDF: $PDFPath" -ForegroundColor Gray
        
        # Call the parser script with -ReturnObject to get the parsed data
        $parsedData = & $parserScript -PDFPath $PDFPath -ReturnObject -PaymentMethod Cash
        
        if ($parsedData) {
            Write-Host "   ✅ Parsing successful!" -ForegroundColor Green
            Write-Host "   📊 Extracted:" -ForegroundColor Cyan
            Write-Host "      - Vendor: $($parsedData.Vendor)" -ForegroundColor White
            Write-Host "      - Invoice: $($parsedData.InvoiceNumber)" -ForegroundColor White
            Write-Host "      - Phones: $($parsedData.ContactInfo.Phone.Count)" -ForegroundColor White
            Write-Host "      - Emails: $($parsedData.ContactInfo.Email.Count)" -ForegroundColor White
            Write-Host "      - Addresses: $($parsedData.PickupAddresses.Count)" -ForegroundColor White
            Write-Host "      - Pickup Dates: $($parsedData.PickupDates.Count)" -ForegroundColor White
            Write-Host "      - Items: $($parsedData.Items.Count)" -ForegroundColor White
            Write-Host "      - Special Notes: $($parsedData.SpecialNotes.Count)" -ForegroundColor White
            
            return $parsedData
        }
        else {
            Write-Warning "Parser returned no data"
            return $null
        }
    }
    catch {
        Write-Warning "Error calling PDF parser: $_"
        Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Yellow
        return $null
    }
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
        Create Date: 01/07/2025
        Version: 1.0.0
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
        Loads template configuration from JSON file
    
    .DESCRIPTION
        Reads a template JSON file and returns the configuration object
        that can be used to populate the GUI form fields.
    
    .PARAMETER TemplatePath
        Full path to the template JSON file
    
    .EXAMPLE
        $config = Import-Template -TemplatePath ".\Templates\Brolyn.json"
    
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 01/07/2025
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
        Create Date: 01/07/2025
        Version: 1.0.0
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
        Create Date: 01/07/2025
        Version: 1.0.0
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
        Creates and displays the main configuration GUI form
    
    .DESCRIPTION
        Builds the complete Windows Forms GUI with all tabs, controls, and event handlers
        for creating and managing logistics email configurations. Includes template management,
        PDF import, and configuration preview features.
    
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 01/07/2025
    #>
    
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Logistics Email Configuration Tool v2.3.2"
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
    $lblTemplateHeader.Text = "Template Manager"
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
    $btnLoadTemplate.Text = "Load Selected Template"
    $btnLoadTemplate.BackColor = [System.Drawing.Color]::LightGreen
    $pnlTemplateManager.Controls.Add($btnLoadTemplate)
    
    $btnDeleteTemplate = New-Object System.Windows.Forms.Button
    $btnDeleteTemplate.Location = New-Object System.Drawing.Point(10, 590)
    $btnDeleteTemplate.Size = New-Object System.Drawing.Size(200, 30)
    $btnDeleteTemplate.Text = "Delete Selected"
    $btnDeleteTemplate.BackColor = [System.Drawing.Color]::LightCoral
    $pnlTemplateManager.Controls.Add($btnDeleteTemplate)
    
    $btnRefreshTemplates = New-Object System.Windows.Forms.Button
    $btnRefreshTemplates.Location = New-Object System.Drawing.Point(10, 625)
    $btnRefreshTemplates.Size = New-Object System.Drawing.Size(200, 30)
    $btnRefreshTemplates.Text = "Refresh Template List"
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
    $btnImportPDF.Size = New-Object System.Drawing.Size(250, 35)
    $btnImportPDF.Text = "Import from PDF Invoice"
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
    $lblSubjectTemplate.Text = "Format: Freight Quote Request - [City, ST ZIP] to [City, ST ZIP] - Pickup [Date]"
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
    $btnGenerateSubject.Text = "Auto-Generate Subject"
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
    $btnSave.Text = "Save Configuration"
    $btnSave.BackColor = [System.Drawing.Color]::LightGreen
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnSave)
    
    $btnSaveTemplate = New-Object System.Windows.Forms.Button
    $btnSaveTemplate.Location = New-Object System.Drawing.Point(910, 700)
    $btnSaveTemplate.Size = New-Object System.Drawing.Size(160, 35)
    $btnSaveTemplate.Text = "Save as Template"
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
    
    function Import-ParsedDataToForm {
        <#
        .SYNOPSIS
            Maps parsed PDF data to form fields - now handles structured address objects
        #>
        param([object]$ParsedData)
        
        if (-not $ParsedData) { return }
        
        # Map vendor
        if ($ParsedData.Vendor) { 
            $cmbAuction.Text = $ParsedData.Vendor 
        }
        
        # Map phone (use first one)
        if ($ParsedData.ContactInfo.Phone -and $ParsedData.ContactInfo.Phone.Count -gt 0) {
            $txtPhone.Text = $ParsedData.ContactInfo.Phone[0]
        }
        
        # Map email (use first one)
        if ($ParsedData.ContactInfo.Email -and $ParsedData.ContactInfo.Email.Count -gt 0) {
            $txtEmail.Text = $ParsedData.ContactInfo.Email[0]
        }
        
        # Map pickup addresses (structured objects with Street, Address2, City, State, Zip)
        if ($ParsedData.PickupAddresses -and $ParsedData.PickupAddresses.Count -gt 0) {
            # Format each address object into multiline text
            $formattedAddresses = @()
            foreach ($addr in $ParsedData.PickupAddresses) {
                if ($addr.Address2) {
                    $formattedAddresses += "$($addr.Street)`r`n$($addr.Address2)`r`n$($addr.City), $($addr.State) $($addr.Zip)"
                } else {
                    $formattedAddresses += "$($addr.Street)`r`n$($addr.City), $($addr.State) $($addr.Zip)"
                }
            }
            
            # If multiple addresses, show selection dialog
            if ($ParsedData.PickupAddresses.Count -gt 1) {
                $selection = Show-AddressSelectionDialog -Addresses $ParsedData.PickupAddresses
                if ($selection) {
                    $txtPickup.Text = $selection
                } else {
                    # User cancelled, use first address
                    $txtPickup.Text = $formattedAddresses[0]
                }
            } else {
                # Single address, just use it
                $txtPickup.Text = $formattedAddresses[0]
            }
        }
        
        # Map pickup dates (use first one or join if multiple)
        if ($ParsedData.PickupDates -and $ParsedData.PickupDates.Count -gt 0) {
            $txtPickupDate.Text = $ParsedData.PickupDates[0]
        }
        
        # Map special notes
        $lstSpecialNotes.Items.Clear()
        if ($ParsedData.SpecialNotes -and $ParsedData.SpecialNotes.Count -gt 0) {
            foreach ($note in $ParsedData.SpecialNotes) {
                $lstSpecialNotes.Items.Add($note) | Out-Null
            }
        }
    }
    
    function Show-AddressSelectionDialog {
        <#
        .SYNOPSIS
            Shows a dialog to select from multiple pickup addresses
        #>
        param([object[]]$Addresses)
        
        $selectionForm = New-Object System.Windows.Forms.Form
        $selectionForm.Text = "Select Pickup Address"
        $selectionForm.Size = New-Object System.Drawing.Size(600, 400)
        $selectionForm.StartPosition = "CenterParent"
        $selectionForm.FormBorderStyle = "FixedDialog"
        $selectionForm.MaximizeBox = $false
        $selectionForm.MinimizeBox = $false
        
        $lblInstruction = New-Object System.Windows.Forms.Label
        $lblInstruction.Location = New-Object System.Drawing.Point(20, 20)
        $lblInstruction.Size = New-Object System.Drawing.Size(560, 40)
        $lblInstruction.Text = "Multiple pickup addresses were found in the invoice.`nPlease select the primary pickup address:"
        $selectionForm.Controls.Add($lblInstruction)
        
        $lstAddresses = New-Object System.Windows.Forms.ListBox
        $lstAddresses.Location = New-Object System.Drawing.Point(20, 70)
        $lstAddresses.Size = New-Object System.Drawing.Size(560, 240)
        $lstAddresses.Font = New-Object System.Drawing.Font("Consolas", 9)
        
        foreach ($addr in $Addresses) {
            if ($addr.Address2) {
                $display = "$($addr.Street) | $($addr.Address2) | $($addr.City), $($addr.State) $($addr.Zip)"
            } else {
                $display = "$($addr.Street) | $($addr.City), $($addr.State) $($addr.Zip)"
            }
            $lstAddresses.Items.Add($display) | Out-Null
        }
        $lstAddresses.SelectedIndex = 0
        $selectionForm.Controls.Add($lstAddresses)
        
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Location = New-Object System.Drawing.Point(400, 320)
        $btnOK.Size = New-Object System.Drawing.Size(80, 30)
        $btnOK.Text = "OK"
        $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $selectionForm.Controls.Add($btnOK)
        
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(490, 320)
        $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
        $btnCancel.Text = "Cancel"
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $selectionForm.Controls.Add($btnCancel)
        
        $selectionForm.AcceptButton = $btnOK
        $selectionForm.CancelButton = $btnCancel
        
        $result = $selectionForm.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $lstAddresses.SelectedIndex -ge 0) {
            $selectedAddr = $Addresses[$lstAddresses.SelectedIndex]
            if ($selectedAddr.Address2) {
                return "$($selectedAddr.Street)`r`n$($selectedAddr.Address2)`r`n$($selectedAddr.City), $($selectedAddr.State) $($selectedAddr.Zip)"
            } else {
                return "$($selectedAddr.Street)`r`n$($selectedAddr.City), $($selectedAddr.State) $($selectedAddr.Zip)"
            }
        }
        
        return $null
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
    
    # Import PDF Button - NOW USES Generic-PDF-Invoice-Parser.ps1
    $btnImportPDF.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
        $openFileDialog.Title = "Select Auction Invoice PDF"
        
        if ($openFileDialog.ShowDialog() -eq "OK") {
            Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
            Write-Host "║   IMPORTING PDF INVOICE DATA                          ║" -ForegroundColor Cyan
            Write-Host "╚════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
            
            # Call the Generic-PDF-Invoice-Parser.ps1 script
            $parsedData = Invoke-PDFInvoiceParser -PDFPath $openFileDialog.FileName
            
            if ($parsedData) {
                # Map the parsed data to the form
                Import-ParsedDataToForm -ParsedData $parsedData
                
                # Build summary message
                $summary = "PDF data imported successfully using Generic-PDF-Invoice-Parser.ps1!`n`n"
                $summary += "Extracted Information:`n"
                $summary += "─────────────────────`n"
                if ($parsedData.Vendor) { $summary += "• Vendor: $($parsedData.Vendor)`n" }
                if ($parsedData.InvoiceNumber) { $summary += "• Invoice: $($parsedData.InvoiceNumber)`n" }
                if ($parsedData.ContactInfo.Phone.Count -gt 0) { 
                    $summary += "• Phones: $($parsedData.ContactInfo.Phone.Count)`n"
                    foreach ($phone in $parsedData.ContactInfo.Phone) {
                        $summary += "  - $phone`n"
                    }
                }
                if ($parsedData.ContactInfo.Email.Count -gt 0) { 
                    $summary += "• Emails: $($parsedData.ContactInfo.Email.Count)`n"
                    foreach ($email in $parsedData.ContactInfo.Email) {
                        $summary += "  - $email`n"
                    }
                }
                if ($parsedData.PickupAddresses.Count -gt 0) { 
                    $summary += "• Pickup Addresses: $($parsedData.PickupAddresses.Count)`n"
                    foreach ($addr in $parsedData.PickupAddresses) {
                        if ($addr.Address2) {
                            $summary += "  - $($addr.Street) ($($addr.Address2)), $($addr.City) $($addr.State)`n"
                        } else {
                            $summary += "  - $($addr.Street), $($addr.City) $($addr.State)`n"
                        }
                    }
                }
                if ($parsedData.PickupDates.Count -gt 0) { 
                    $summary += "• Pickup Dates: $($parsedData.PickupDates.Count)`n"
                    foreach ($date in $parsedData.PickupDates) {
                        $summary += "  - $date`n"
                    }
                }
                if ($parsedData.Items.Count -gt 0) { $summary += "• Items: $($parsedData.Items.Count)`n" }
                if ($parsedData.SpecialNotes.Count -gt 0) { $summary += "• Special Notes: $($parsedData.SpecialNotes.Count)`n" }
                $summary += "`nPlease review and adjust as needed."
                
                [System.Windows.Forms.MessageBox]::Show(
                    $summary,
                    "Import Successful",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Could not parse PDF invoice.`n`nPossible issues:`n• Generic-PDF-Invoice-Parser.ps1 not found in script directory`n• PDF extraction failed`n• pdftotext not installed`n`nPlease check console output for details.",
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
            $pickupText = $txtPickup.Text
            $pickupCity = "Unknown"
            $pickupState = "XX"
            $pickupZip = "00000"
            
            # Handle multiline address format (split and look at last line for city/state/zip)
            $pickupLines = $pickupText -split "`r`n|`r|`n"
            $cityStateLine = $pickupLines | Where-Object { $_ -match '([A-Za-z\s]+),?\s+([A-Z]{2})\s+(\d{5})' } | Select-Object -Last 1
            
            if ($cityStateLine) {
                # Extract city, state, and ZIP from the line
                if ($cityStateLine -match '([A-Za-z][A-Za-z\s.''()-]+),?\s+([A-Z]{2})\s+(\d{5})') {
                    $pickupCity = $Matches[1].Trim() -replace '\s+', ' '
                    $pickupState = $Matches[2].Trim()
                    $pickupZip = $Matches[3].Trim()
                }
            }
            # Fallback: try single-line formats
            elseif ($pickupText -match ',\s*([A-Za-z\s]+),\s*([A-Z]{2})\s+(\d{5})') {
                $pickupCity = $Matches[1].Trim()
                $pickupState = $Matches[2].Trim()
                $pickupZip = $Matches[3].Trim()
            }
            elseif ($pickupText -match '\(([A-Za-z\s]+)\).*?([A-Z]{2})\s+(\d{5})') {
                $pickupCity = $Matches[1].Trim()
                $pickupState = $Matches[2].Trim()
                $pickupZip = $Matches[3].Trim()
            }
            
            $deliveryText = $cmbDelivery.Text
            $deliveryCity = "Unknown"
            $deliveryState = "XX"
            $deliveryZip = "00000"
            
            # Handle delivery address (usually single-line)
            if ($deliveryText -match ',\s*([A-Za-z\s]+),\s*([A-Z]{2})\s+(\d{5})') {
                $deliveryCity = $Matches[1].Trim()
                $deliveryState = $Matches[2].Trim()
                $deliveryZip = $Matches[3].Trim()
            }
            
            $pickupDateText = $txtPickupDate.Text
            $pickupDate = "TBD"
            
            # Extract date from pickup datetime - handle multiple formats
            # Format 1: Full date with year (MM/DD/YYYY)
            if ($pickupDateText -match '(\d{1,2}/\d{1,2}/\d{4})') {
                $pickupDate = $Matches[1]
            }
            # Format 2: Date range with "thru" or "through" - extract first date
            elseif ($pickupDateText -match '(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+(\d{1,2}/\d{1,2})(?:/(\d{2,4}))?\s+(?:thru|through|to|-)\s+') {
                $month = $Matches[1].Split('/')[0]
                $day = $Matches[1].Split('/')[1]
                $year = if ($Matches[2]) { $Matches[2] } else { (Get-Date).Year }
                # Handle 2-digit year
                if ($year.Length -eq 2) { $year = "20$year" }
                $pickupDate = "$month/$day/$year"
            }
            # Format 3: Simple date range without day names (10/7 thru 10/10)
            elseif ($pickupDateText -match '(\d{1,2}/\d{1,2})(?:/(\d{2,4}))?\s+(?:thru|through|to|-)\s+') {
                $month = $Matches[1].Split('/')[0]
                $day = $Matches[1].Split('/')[1]
                $year = if ($Matches[2]) { $Matches[2] } else { (Get-Date).Year }
                if ($year.Length -eq 2) { $year = "20$year" }
                $pickupDate = "$month/$day/$year"
            }
            # Format 4: Month name with year (October 10th, 2025)
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
            # Format 5: Single date without year (10/7)
            elseif ($pickupDateText -match '(\d{1,2}/\d{1,2})(?!\d)') {
                $month = $Matches[1].Split('/')[0]
                $day = $Matches[1].Split('/')[1]
                $year = (Get-Date).Year
                $pickupDate = "$month/$day/$year"
            }
            
            # Build subject with ZIP codes for professional freight industry standard
            $subject = "Freight Quote Request - $pickupCity, $pickupState $pickupZip to $deliveryCity, $deliveryState $deliveryZip - Pickup $pickupDate"
            $txtSubject.Text = $subject
            
            [System.Windows.Forms.MessageBox]::Show(
                "Subject line generated!`n`nSubject:`n$subject`n`n✓ Includes ZIP codes for accurate freight routing`n✓ Ready for professional logistics quotes`n`nPlease review and edit if needed.",
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
Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   LOGISTICS EMAIL CONFIGURATION TOOL v2.3.2           ║" -ForegroundColor Cyan
Write-Host "║   with Generic-PDF-Invoice-Parser Integration        ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

# Auto-load PDF if provided
if ($PDFInvoice -and (Test-Path $PDFInvoice)) {
    Write-Host "Auto-loading PDF: $PDFInvoice" -ForegroundColor Yellow
    $script:AutoLoadPDF = $PDFInvoice
}

New-ConfigurationGUI
#endregion