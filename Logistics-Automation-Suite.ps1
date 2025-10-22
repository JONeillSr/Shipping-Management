<#
.SYNOPSIS
    Logistics Automation Suite - Enhanced Master Launcher v2.4.1

.DESCRIPTION
    Central hub with analytics dashboard, enhanced carrier tracking,
    universal PDF invoice parser integration, and streamlined email generation.
    Includes automatic PDF conversion with Foxit/Edge/Chrome support.

.EXAMPLE
    .\Logistics-Automation-Suite.ps1

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 01/08/2025
    Version: 2.4.1
    Change Date: 10/17/2025
    Change Purpose: Fixed HTML generation parsing errors with proper here-string usage

.CHANGELOG
    2.4.1 - 10/17/2025 - Fixed PowerShell parsing errors in HTML generation
                       - Properly escaped all HTML tags using here-strings
                       - Resolved '<' operator and onclick parsing issues
                       - Improved error handling and console output
    2.4.0 - 10/10/2025 - Integrated Convert-HTMLtoPDF.ps1 helper script
                       - Automatic PDF conversion using Foxit, Edge, or Chrome
                       - Fast, reliable PDF creation (3-5 seconds vs overnight hangs)
                       - No more Word COM automation issues
    2.3.0 - 10/10/2025 - Removed Word PDF conversion from default workflow
                       - Prevents script from hanging overnight on Word COM issues
                       - Users get clear instructions to print HTML to PDF manually
    2.2.1 - 10/09/2025 - Fixed maxImages variable initialization with InputBox prompt
                       - Added validation for 1-10 range with default of 3
                       - Ensures parameter always has a value when passed to script
    2.2.0 - 10/09/2025 - Added MaxImagesPerLot parameter to GUI workflow
                       - Users can now choose 1-10 images per lot (default 3)
                       - Added Microsoft.VisualBasic assembly for InputBox
    2.1.1 - 10/09/2025 - Fixed apostrophe/single quote handling in file paths
    2.1.0 - 10/09/2025 - Removed confusing mode selection prompt
    2.0.3 - 10/09/2025 - Config file selection now available in Standard mode
    2.0.2 - 10/09/2025 - Fixed execution policy issue with script execution
    2.0.1 - 10/09/2025 - Fixed Generate Email button error handling
    2.0.0 - 01/08/2025 - Added Analytics Dashboard & PDF Parser Integration
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$script:ScriptDirectory = $PSScriptRoot

#region Helper Functions

function Invoke-ScriptWithBypass {
    <#
    .SYNOPSIS
        Executes a PowerShell script with execution policy bypass

    .DESCRIPTION
        Helper function to execute PowerShell scripts with bypass policy
        while properly capturing output and handling parameters

    .PARAMETER ScriptPath
        Full path to the PowerShell script to execute

    .PARAMETER Parameters
        Hashtable of parameters to pass to the script

    .EXAMPLE
        $params = @{
            CSVPath = "C:\data.csv"
            Verbose = $true
        }
        Invoke-ScriptWithBypass -ScriptPath "C:\script.ps1" -Parameters $params

    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 10/09/2025
        Version: 1.0.0
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ScriptPath,

        [Parameter(Mandatory=$false)]
        [hashtable]$Parameters = @{}
    )

    # Build parameter string
    $paramString = ""
    foreach ($key in $Parameters.Keys) {
        $value = $Parameters[$key]
        if ($value -is [bool]) {
            if ($value) {
                $paramString += " -$key"
            }
        }
        else {
            $paramString += " -$key `"$value`""
        }
    }

    # Execute with bypass
    $command = "& '$ScriptPath' $paramString"

    try {
        $result = powershell.exe -ExecutionPolicy Bypass -Command $command
        return $result
    }
    catch {
        throw
    }
}

function Test-ScriptAvailability {
    <#
    .SYNOPSIS
        Checks if a required script file exists

    .DESCRIPTION
        Validates that a script file exists in the expected location
        and returns detailed information about its status

    .PARAMETER ScriptName
        Name of the script file to check

    .PARAMETER Required
        Whether the script is required for core functionality

    .EXAMPLE
        Test-ScriptAvailability -ScriptName "Generate-Email.ps1" -Required $true

    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 10/17/2025
        Version: 1.0.0
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ScriptName,

        [Parameter(Mandatory=$false)]
        [bool]$Required = $false
    )

    $scriptPath = Join-Path $script:ScriptDirectory $ScriptName

    return @{
        Name = $ScriptName
        Path = $scriptPath
        Exists = (Test-Path $scriptPath)
        Required = $Required
    }
}

#endregion

#region Main Menu GUI

function Show-EnhancedMainMenu {
    <#
    .SYNOPSIS
        Displays the main menu GUI for the Logistics Automation Suite

    .DESCRIPTION
        Creates and displays a comprehensive GUI with all available tools,
        quick actions, and utilities for logistics automation workflow

    .EXAMPLE
        Show-EnhancedMainMenu

    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 01/08/2025
        Version: 2.4.1
        Change Date: 10/17/2025
        Change Purpose: Updated version display and help text
    #>

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Logistics Automation Suite v2.4.1 - JT Custom Trailers"
    $form.Size = New-Object System.Drawing.Size(950, 750)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::White

    #region Header Section

    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Location = New-Object System.Drawing.Point(0, 0)
    $pnlHeader.Size = New-Object System.Drawing.Size(950, 100)
    $pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(44, 62, 80)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(900, 35)
    $lblTitle.Text = "Logistics Automation Suite v2.4.1"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 20, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $pnlHeader.Controls.Add($lblTitle)

    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(20, 55)
    $lblSubtitle.Size = New-Object System.Drawing.Size(900, 20)
    $lblSubtitle.Text = "Complete freight quote automation with analytics & performance tracking"
    $lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblSubtitle.ForeColor = [System.Drawing.Color]::LightGray
    $pnlHeader.Controls.Add($lblSubtitle)

    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.Location = New-Object System.Drawing.Point(20, 75)
    $lblVersion.Size = New-Object System.Drawing.Size(900, 18)
    $lblVersion.Text = "FIXED: HTML parsing errors • Enhanced error handling • Improved console output"
    $lblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblVersion.ForeColor = [System.Drawing.Color]::FromArgb(46, 204, 113)
    $pnlHeader.Controls.Add($lblVersion)

    $form.Controls.Add($pnlHeader)

    #endregion

    #region Main Tools Section

    $yPos = 120

    $lblMainTools = New-Object System.Windows.Forms.Label
    $lblMainTools.Location = New-Object System.Drawing.Point(30, $yPos)
    $lblMainTools.Size = New-Object System.Drawing.Size(400, 25)
    $lblMainTools.Text = "Main Tools"
    $lblMainTools.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblMainTools)

    $yPos += 40

    # Row 1 - Configuration, PDF Parser, Recipients
    $btnConfig = New-Object System.Windows.Forms.Button
    $btnConfig.Location = New-Object System.Drawing.Point(30, $yPos)
    $btnConfig.Size = New-Object System.Drawing.Size(280, 60)
    $btnConfig.Text = "Configuration Tool`n   Create & import auction configs"
    $btnConfig.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $btnConfig.ForeColor = [System.Drawing.Color]::White
    $btnConfig.FlatStyle = "Flat"
    $btnConfig.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnConfig.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($btnConfig)

    $btnPDFParser = New-Object System.Windows.Forms.Button
    $btnPDFParser.Location = New-Object System.Drawing.Point(330, $yPos)
    $btnPDFParser.Size = New-Object System.Drawing.Size(280, 60)
    $btnPDFParser.Text = "PDF Invoice Parser`n   Universal multi-vendor parser"
    $btnPDFParser.BackColor = [System.Drawing.Color]::FromArgb(41, 128, 185)
    $btnPDFParser.ForeColor = [System.Drawing.Color]::White
    $btnPDFParser.FlatStyle = "Flat"
    $btnPDFParser.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnPDFParser.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($btnPDFParser)

    $btnRecipients = New-Object System.Windows.Forms.Button
    $btnRecipients.Location = New-Object System.Drawing.Point(630, $yPos)
    $btnRecipients.Size = New-Object System.Drawing.Size(280, 60)
    $btnRecipients.Text = "Recipient Manager`n   Manage carrier contacts"
    $btnRecipients.BackColor = [System.Drawing.Color]::FromArgb(46, 204, 113)
    $btnRecipients.ForeColor = [System.Drawing.Color]::White
    $btnRecipients.FlatStyle = "Flat"
    $btnRecipients.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnRecipients.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($btnRecipients)

    $yPos += 80

    # Row 2 - Generate Email, Quote Tracker, Analytics
    $btnGenerate = New-Object System.Windows.Forms.Button
    $btnGenerate.Location = New-Object System.Drawing.Point(30, $yPos)
    $btnGenerate.Size = New-Object System.Drawing.Size(280, 60)
    $btnGenerate.Text = "Generate Email`n   Create quote requests"
    $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(155, 89, 182)
    $btnGenerate.ForeColor = [System.Drawing.Color]::White
    $btnGenerate.FlatStyle = "Flat"
    $btnGenerate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnGenerate.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($btnGenerate)

    $btnTracker = New-Object System.Windows.Forms.Button
    $btnTracker.Location = New-Object System.Drawing.Point(330, $yPos)
    $btnTracker.Size = New-Object System.Drawing.Size(280, 60)
    $btnTracker.Text = "Quote Tracker`n   Track quotes & costs"
    $btnTracker.BackColor = [System.Drawing.Color]::FromArgb(230, 126, 34)
    $btnTracker.ForeColor = [System.Drawing.Color]::White
    $btnTracker.FlatStyle = "Flat"
    $btnTracker.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnTracker.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($btnTracker)

    $btnAnalytics = New-Object System.Windows.Forms.Button
    $btnAnalytics.Location = New-Object System.Drawing.Point(630, $yPos)
    $btnAnalytics.Size = New-Object System.Drawing.Size(280, 60)
    $btnAnalytics.Text = "Analytics Dashboard`n   Cost trends & performance"
    $btnAnalytics.BackColor = [System.Drawing.Color]::FromArgb(231, 76, 60)
    $btnAnalytics.ForeColor = [System.Drawing.Color]::White
    $btnAnalytics.FlatStyle = "Flat"
    $btnAnalytics.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnAnalytics.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($btnAnalytics)

    $yPos += 100

    #endregion

    #region Quick Actions Section

    $lblQuickActions = New-Object System.Windows.Forms.Label
    $lblQuickActions.Location = New-Object System.Drawing.Point(30, $yPos)
    $lblQuickActions.Size = New-Object System.Drawing.Size(400, 25)
    $lblQuickActions.Text = "Quick Actions & Workflows"
    $lblQuickActions.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblQuickActions)

    $yPos += 40

    # Quick Workflow Buttons
    $btnQuickBrolyn = New-Object System.Windows.Forms.Button
    $btnQuickBrolyn.Location = New-Object System.Drawing.Point(30, $yPos)
    $btnQuickBrolyn.Size = New-Object System.Drawing.Size(200, 45)
    $btnQuickBrolyn.Text = "New Brolyn Auction"
    $btnQuickBrolyn.BackColor = [System.Drawing.Color]::LightGreen
    $btnQuickBrolyn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnQuickBrolyn)

    $btnQuickGeneric = New-Object System.Windows.Forms.Button
    $btnQuickGeneric.Location = New-Object System.Drawing.Point(240, $yPos)
    $btnQuickGeneric.Size = New-Object System.Drawing.Size(200, 45)
    $btnQuickGeneric.Text = "New Generic Auction"
    $btnQuickGeneric.BackColor = [System.Drawing.Color]::LightBlue
    $btnQuickGeneric.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnQuickGeneric)

    $btnParseInvoice = New-Object System.Windows.Forms.Button
    $btnParseInvoice.Location = New-Object System.Drawing.Point(450, $yPos)
    $btnParseInvoice.Size = New-Object System.Drawing.Size(200, 45)
    $btnParseInvoice.Text = "Parse PDF Invoice"
    $btnParseInvoice.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $btnParseInvoice.ForeColor = [System.Drawing.Color]::White
    $btnParseInvoice.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnParseInvoice)

    $btnTemplates = New-Object System.Windows.Forms.Button
    $btnTemplates.Location = New-Object System.Drawing.Point(660, $yPos)
    $btnTemplates.Size = New-Object System.Drawing.Size(250, 45)
    $btnTemplates.Text = "Create Starter Templates"
    $btnTemplates.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($btnTemplates)

    $yPos += 60

    #endregion

    #region Utilities Section

    $lblUtilities = New-Object System.Windows.Forms.Label
    $lblUtilities.Location = New-Object System.Drawing.Point(30, $yPos)
    $lblUtilities.Size = New-Object System.Drawing.Size(400, 25)
    $lblUtilities.Text = "Utilities"
    $lblUtilities.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblUtilities)

    $yPos += 40

    $btnExcelConverter = New-Object System.Windows.Forms.Button
    $btnExcelConverter.Location = New-Object System.Drawing.Point(30, $yPos)
    $btnExcelConverter.Size = New-Object System.Drawing.Size(200, 45)
    $btnExcelConverter.Text = "Excel to CSV Converter"
    $btnExcelConverter.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($btnExcelConverter)

    $btnViewAnalytics = New-Object System.Windows.Forms.Button
    $btnViewAnalytics.Location = New-Object System.Drawing.Point(240, $yPos)
    $btnViewAnalytics.Size = New-Object System.Drawing.Size(200, 45)
    $btnViewAnalytics.Text = "View Analytics Report"
    $btnViewAnalytics.BackColor = [System.Drawing.Color]::FromArgb(241, 196, 15)
    $btnViewAnalytics.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnViewAnalytics)

    $btnDataBackup = New-Object System.Windows.Forms.Button
    $btnDataBackup.Location = New-Object System.Drawing.Point(450, $yPos)
    $btnDataBackup.Size = New-Object System.Drawing.Size(200, 45)
    $btnDataBackup.Text = "Backup Data Files"
    $btnDataBackup.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.Controls.Add($btnDataBackup)

    $yPos += 70

    #endregion

    #region Help & Documentation

    $grpHelp = New-Object System.Windows.Forms.GroupBox
    $grpHelp.Location = New-Object System.Drawing.Point(30, $yPos)
    $grpHelp.Size = New-Object System.Drawing.Size(880, 85)
    $grpHelp.Text = "Getting Started"

    $lblHelp = New-Object System.Windows.Forms.Label
    $lblHelp.Location = New-Object System.Drawing.Point(15, 25)
    $lblHelp.Size = New-Object System.Drawing.Size(850, 50)
    $lblHelp.Text = @"
1. Use "PDF Invoice Parser" to extract data from auction invoices
2. Or use "Configuration Tool" to manually set up auction details
3. Manage freight companies in "Recipient Manager"
4. Click "Generate Email" to create quote requests (HTML files created instantly, print to PDF in browser!)
5. Track responses and analyze costs in "Quote Tracker" and "Analytics Dashboard"
"@
    $lblHelp.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblHelp.ForeColor = [System.Drawing.Color]::FromArgb(52, 73, 94)
    $grpHelp.Controls.Add($lblHelp)

    $form.Controls.Add($grpHelp)

    #endregion

    #region Footer

    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Location = New-Object System.Drawing.Point(790, 670)
    $btnExit.Size = New-Object System.Drawing.Size(120, 35)
    $btnExit.Text = "Exit"
    $btnExit.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($btnExit)

    #endregion

    #region Event Handlers

    $btnConfig.Add_Click({
        <#
        .SYNOPSIS
            Opens Configuration Tool
        .NOTES
            Launches the Logistics Configuration GUI for creating and managing
            auction configuration files
        #>
        $configScript = Join-Path $script:ScriptDirectory "Logistics-Config-GUI.ps1"
        if (Test-Path $configScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$configScript`""
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Configuration script not found:`n$configScript",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnPDFParser.Add_Click({
        <#
        .SYNOPSIS
            Opens PDF Invoice Parser
        .NOTES
            Launches the universal PDF invoice parser for extracting
            auction data from vendor PDFs
        #>
        $parserScript = Join-Path $script:ScriptDirectory "Generic-PDF-Invoice-Parser.ps1"
        if (Test-Path $parserScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$parserScript`" -GUI"
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "PDF Parser script not found:`n$parserScript",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnRecipients.Add_Click({
        <#
        .SYNOPSIS
            Opens Recipient Manager
        .NOTES
            Launches the freight recipient manager for maintaining
            carrier contact information
        #>
        $recipientScript = Join-Path $script:ScriptDirectory "Freight-Recipient-Manager.ps1"
        if (Test-Path $recipientScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$recipientScript`""
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Recipient Manager script not found:`n$recipientScript",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnGenerate.Add_Click({
        <#
        .SYNOPSIS
            Generate Email button - Main workflow for creating freight quote requests
        .DESCRIPTION
            Guides user through selecting CSV, images, and configuration,
            then generates HTML report with embedded images
        .NOTES
            Author: John O'Neill Sr.
            Company: Azure Innovators
            Version: 2.4.1
            Change Date: 10/17/2025
            Change Purpose: Enhanced error handling and user feedback
        #>

        try {
            Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
            Write-Host "║    Starting Email Generation Process                  ║" -ForegroundColor Cyan
            Write-Host "╚════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

            # Step 1: Select CSV file
            Write-Host "Step 1: Selecting CSV file..." -ForegroundColor Yellow
            $openCSV = New-Object System.Windows.Forms.OpenFileDialog
            $openCSV.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
            $openCSV.Title = "Select Auction Lots CSV File"
            $openCSV.InitialDirectory = [Environment]::GetFolderPath('Desktop')

            if ($openCSV.ShowDialog() -ne "OK") {
                Write-Host "CSV selection cancelled by user" -ForegroundColor Yellow
                return
            }

            $csvPath = $openCSV.FileName
            Write-Host "   ✓ CSV selected: $(Split-Path $csvPath -Leaf)" -ForegroundColor Green

            # Step 2: Select Image Directory
            Write-Host "`nStep 2: Selecting image directory..." -ForegroundColor Yellow
            $imageDir = New-Object System.Windows.Forms.FolderBrowserDialog
            $imageDir.Description = "Select Directory Containing Lot Images"
            $imageDir.RootFolder = [System.Environment+SpecialFolder]::Desktop

            if ($imageDir.ShowDialog() -ne "OK") {
                Write-Host "Image directory selection cancelled by user" -ForegroundColor Yellow
                return
            }

            $imagePath = $imageDir.SelectedPath
            Write-Host "   ✓ Images directory: $(Split-Path $imagePath -Leaf)" -ForegroundColor Green

            # Step 3: Ask for maximum images per lot
            Write-Host "`nStep 3: Maximum images per lot..." -ForegroundColor Yellow

            $maxImages = 3  # Default value

            $imageCountInput = [Microsoft.VisualBasic.Interaction]::InputBox(
                "How many images per lot would you like to include?`n`n" +
                "Enter a number between 1 and 10.`n" +
                "Default is 3 images per lot.`n`n" +
                "More images = larger HTML files but better detail.",
                "Maximum Images Per Lot",
                "3"
            )

            if ($imageCountInput -ne "") {
                # Validate and parse the input
                $parsedValue = 0
                if ([int]::TryParse($imageCountInput, [ref]$parsedValue)) {
                    if ($parsedValue -ge 1 -and $parsedValue -le 10) {
                        $maxImages = $parsedValue
                        Write-Host "   ✓ Using $maxImages images per lot" -ForegroundColor Green
                    }
                    else {
                        Write-Host "   ⚠ Invalid range ($parsedValue). Using default (3)" -ForegroundColor Yellow
                        $maxImages = 3
                    }
                }
                else {
                    Write-Host "   ⚠ Invalid input. Using default (3)" -ForegroundColor Yellow
                    $maxImages = 3
                }
            }
            else {
                Write-Host "   ✓ Using default (3 images per lot)" -ForegroundColor Green
                $maxImages = 3
            }

            # Step 4: Ask if user wants to use a config file
            $configPath = $null

            Write-Host "`nStep 4: Configuration file..." -ForegroundColor Yellow
            $useConfig = [System.Windows.Forms.MessageBox]::Show(
                "Do you have a configuration JSON file with auction details?`n`n" +
                "Config files contain pickup address, delivery info, special notes, etc.`n`n" +
                "YES = Select config file (recommended)`n" +
                "NO = Use basic format without config",
                "Use Configuration File?",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($useConfig -eq 'Yes') {
                $openConfig = New-Object System.Windows.Forms.OpenFileDialog
                $openConfig.Filter = "JSON Configuration (*.json)|*.json|All files (*.*)|*.*"
                $openConfig.Title = "Select Configuration JSON File"
                $openConfig.InitialDirectory = Join-Path $script:ScriptDirectory "Templates"

                if ($openConfig.ShowDialog() -eq "OK") {
                    $configPath = $openConfig.FileName
                    Write-Host "   ✓ Config selected: $(Split-Path $configPath -Leaf)" -ForegroundColor Green
                }
                else {
                    Write-Host "   ℹ No config selected, continuing without config" -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "   ℹ Skipping config file" -ForegroundColor Yellow
            }

            # Step 5: Verify Generate-LogisticsEmail.ps1 exists
            Write-Host "`nStep 5: Verifying email generation script..." -ForegroundColor Yellow
            $genScript = Join-Path $script:ScriptDirectory "Generate-LogisticsEmail.ps1"

            if (!(Test-Path $genScript)) {
                $errorMsg = "Generate-LogisticsEmail.ps1 not found at:`n$genScript"
                Write-Host "   ✗ ERROR: $errorMsg" -ForegroundColor Red
                [System.Windows.Forms.MessageBox]::Show(
                    $errorMsg,
                    "Script Not Found",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
            Write-Host "   ✓ Email generation script found" -ForegroundColor Green

            Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Green
            Write-Host "║    Executing Email Generation Script                  ║" -ForegroundColor Green
            Write-Host "╚════════════════════════════════════════════════════════╝`n" -ForegroundColor Green

            # Step 6: Build command string for execution
            # Escape single quotes/apostrophes in paths for PowerShell
            $escapedGenScript = $genScript -replace "'", "''"
            $escapedCsvPath = $csvPath -replace "'", "''"
            $escapedImagePath = $imagePath -replace "'", "''"

            $command = "& '$escapedGenScript' -CSVPath '$escapedCsvPath' -ImageDirectory '$escapedImagePath' -MaxImagesPerLot $maxImages -CreateOutlookDraft -ShowDashboard"

            # Add config path if provided
            if ($configPath) {
                $escapedConfigPath = $configPath -replace "'", "''"
                $command += " -ConfigPath '$escapedConfigPath'"
                Write-Host "Configuration:" -ForegroundColor Cyan
                Write-Host "   Config: $(Split-Path $configPath -Leaf)" -ForegroundColor Gray
            }
            else {
                Write-Host "Configuration:" -ForegroundColor Cyan
                Write-Host "   Running without config file (basic mode)" -ForegroundColor Gray
            }

            Write-Host "`nParameters:" -ForegroundColor Cyan
            Write-Host "   CSV: $(Split-Path $csvPath -Leaf)" -ForegroundColor Gray
            Write-Host "   Images: $(Split-Path $imagePath -Leaf)" -ForegroundColor Gray
            Write-Host "   Max Images Per Lot: $maxImages" -ForegroundColor Gray
            Write-Host "`n" + ("─" * 60) -ForegroundColor Gray
            Write-Host "Processing... Please wait..." -ForegroundColor Yellow
            Write-Host ("─" * 60) + "`n" -ForegroundColor Gray

            # Step 7: Execute with execution policy bypass
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = "powershell.exe"
            $psi.Arguments = "-ExecutionPolicy Bypass -Command `"$command`""
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $false

            $process = [System.Diagnostics.Process]::Start($psi)
            $process.WaitForExit()

            $exitCode = $process.ExitCode

            Write-Host "`n" + ("─" * 60) -ForegroundColor Gray

            if ($exitCode -eq 0) {
                Write-Host "╔════════════════════════════════════════════════════════╗" -ForegroundColor Green
                Write-Host "║          EMAIL GENERATED SUCCESSFULLY! ✓               ║" -ForegroundColor Green
                Write-Host "╚════════════════════════════════════════════════════════╝" -ForegroundColor Green

                [System.Windows.Forms.MessageBox]::Show(
                    "Email generation completed successfully!`n`n" +
                    "✓ HTML report created with embedded images`n" +
                    "✓ Check the Output directory for files`n" +
                    "✓ Open HTML in browser and print to PDF`n`n" +
                    "The dashboard should display with statistics.",
                    "Success",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            else {
                Write-Host "⚠ Script exited with code: $exitCode" -ForegroundColor Yellow
                Write-Host "Check the console output above for details." -ForegroundColor Gray

                [System.Windows.Forms.MessageBox]::Show(
                    "Script completed with exit code: $exitCode`n`n" +
                    "Please check the console window for details.`n" +
                    "Review the Output and Logs directories for more information.",
                    "Process Complete",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
            }
        }
        catch {
            # Catch any unexpected errors
            $errorDetails = @"
╔════════════════════════════════════════════════════════╗
║               ERROR OCCURRED ✗                         ║
╚════════════════════════════════════════════════════════╝

ERROR MESSAGE:
$($_.Exception.Message)

ERROR TYPE: $($_.Exception.GetType().FullName)
CATEGORY: $($_.CategoryInfo.Category)
TARGET: $($_.TargetObject)

STACK TRACE:
$($_.ScriptStackTrace)
"@

            Write-Host $errorDetails -ForegroundColor Red

            # Save error to log file
            try {
                $errorLogDir = Join-Path $script:ScriptDirectory "Logs"
                if (!(Test-Path $errorLogDir)) {
                    New-Item -ItemType Directory -Path $errorLogDir -Force | Out-Null
                }

                $errorLogPath = Join-Path $errorLogDir ("GUI_Error_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")
                $errorDetails | Out-File -FilePath $errorLogPath -Encoding UTF8

                Write-Host "`nError details saved to: $errorLogPath" -ForegroundColor Yellow
            }
            catch {
                Write-Host "`nCould not save error log: $($_.Exception.Message)" -ForegroundColor Red
            }

            [System.Windows.Forms.MessageBox]::Show(
                "An error occurred during email generation:`n`n" +
                "$($_.Exception.Message)`n`n" +
                "Error details have been logged for troubleshooting.",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        finally {
            Write-Host "`n" + ("═" * 60) -ForegroundColor Cyan
            Write-Host "Operation Complete - " + (Get-Date -Format "HH:mm:ss") -ForegroundColor Cyan
            Write-Host ("═" * 60) + "`n" -ForegroundColor Cyan
        }
    })

    $btnTracker.Add_Click({
        <#
        .SYNOPSIS
            Opens Quote Tracker
        .NOTES
            Launches the auction quote tracker for managing and tracking
            freight quotes and carrier responses
        #>
        $trackerScript = Join-Path $script:ScriptDirectory "Auction-Quote-Tracker.ps1"
        if (Test-Path $trackerScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$trackerScript`""
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Quote Tracker script not found:`n$trackerScript",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnAnalytics.Add_Click({
        <#
        .SYNOPSIS
            Opens Analytics Dashboard
        .NOTES
            Launches the analytics dashboard for viewing cost trends,
            carrier performance, and logistics statistics
        #>
        $analyticsScript = Join-Path $script:ScriptDirectory "Logistics-Analytics-Dashboard.ps1"
        if (Test-Path $analyticsScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$analyticsScript`""
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Analytics Dashboard script not found:`n$analyticsScript`n`nPlease ensure the analytics script is in the suite directory.",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnQuickBrolyn.Add_Click({
        <#
        .SYNOPSIS
            Quick launch for Brolyn Auction configuration
        .NOTES
            Launches configuration tool with Brolyn template pre-loaded
        #>
        $configScript = Join-Path $script:ScriptDirectory "Logistics-Config-GUI.ps1"
        if (Test-Path $configScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$configScript`" -LoadTemplate `"Brolyn_Auctions`""
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Configuration script not found",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnQuickGeneric.Add_Click({
        <#
        .SYNOPSIS
            Quick launch for Generic Auction configuration
        .NOTES
            Launches configuration tool with generic template pre-loaded
        #>
        $configScript = Join-Path $script:ScriptDirectory "Logistics-Config-GUI.ps1"
        if (Test-Path $configScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$configScript`" -LoadTemplate `"Generic_Template`""
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Configuration script not found",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnParseInvoice.Add_Click({
        <#
        .SYNOPSIS
            Quick PDF Invoice parser with file picker
        .NOTES
            Opens file dialog and launches parser with selected PDF
        #>
        $openPDF = New-Object System.Windows.Forms.OpenFileDialog
        $openPDF.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
        $openPDF.Title = "Select Invoice PDF to Parse"
        $openPDF.InitialDirectory = [Environment]::GetFolderPath('Desktop')

        if ($openPDF.ShowDialog() -eq "OK") {
            $parserScript = Join-Path $script:ScriptDirectory "Generic-PDF-Invoice-Parser.ps1"
            if (Test-Path $parserScript) {
                Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$parserScript`" -PDFPath `"$($openPDF.FileName)`" -OutputFormat Config"
            }
            else {
                [System.Windows.Forms.MessageBox]::Show(
                    "PDF Parser script not found",
                    "Script Not Found",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        }
    })

    $btnTemplates.Add_Click({
        <#
        .SYNOPSIS
            Creates starter templates
        .NOTES
            Launches template creation script to generate starter
            configuration files for common auction types
        #>
        $templateScript = Join-Path $script:ScriptDirectory "Create-StarterTemplates.ps1"
        if (Test-Path $templateScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$templateScript`""
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Template creation script not found:`n$templateScript",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnExcelConverter.Add_Click({
        <#
        .SYNOPSIS
            Opens Excel to CSV converter
        .NOTES
            Launches Excel to CSV conversion utility for processing
            auction data from spreadsheets
        #>
        $excelScript = Join-Path $script:ScriptDirectory "Export-ExcelToCSV.ps1"
        if (Test-Path $excelScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$excelScript`" -InteractiveMode"
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Excel converter script not found:`n$excelScript",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnViewAnalytics.Add_Click({
        <#
        .SYNOPSIS
            Opens Analytics Dashboard (utility shortcut)
        .NOTES
            Same as main Analytics button - provides quick access
            from utilities section
        #>
        $analyticsScript = Join-Path $script:ScriptDirectory "Logistics-Analytics-Dashboard.ps1"
        if (Test-Path $analyticsScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$analyticsScript`""
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Analytics Dashboard script not found",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnDataBackup.Add_Click({
        <#
        .SYNOPSIS
            Backs up all data files to timestamped folder
        .DESCRIPTION
            Creates a backup of all JSON data files and templates
            to a timestamped backup directory
        .NOTES
            Author: John O'Neill Sr.
            Company: Azure Innovators
            Version: 1.0.0
        #>
        try {
            $backupDir = Join-Path $script:ScriptDirectory ("Backups\Backup_" + (Get-Date -Format 'yyyyMMdd_HHmmss'))

            # Create Backups directory if it doesn't exist
            $backupsRoot = Join-Path $script:ScriptDirectory "Backups"
            if (!(Test-Path $backupsRoot)) {
                New-Item -ItemType Directory -Path $backupsRoot -Force | Out-Null
            }

            # Create timestamped backup directory
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

            # Backup data files
            $dataDir = Join-Path $script:ScriptDirectory "Data"
            $dataFiles = @(
                "AuctionQuotes.json",
                "FreightRecipients.json",
                "InvoicePatterns.json"
            )

            $backedUp = 0
            foreach ($fileName in $dataFiles) {
                $filePath = Join-Path $dataDir $fileName
                if (Test-Path $filePath) {
                    Copy-Item $filePath -Destination $backupDir -Force
                    $backedUp++
                }
            }

            # Backup templates directory
            $templatesDir = Join-Path $script:ScriptDirectory "Templates"
            if (Test-Path $templatesDir) {
                $templateBackupDir = Join-Path $backupDir "Templates"
                Copy-Item $templatesDir -Destination $templateBackupDir -Force -Recurse
            }

            [System.Windows.Forms.MessageBox]::Show(
                "Backup completed successfully!`n`n" +
                "✓ Backed up $backedUp data files`n" +
                "✓ Backed up templates directory`n`n" +
                "Location: $backupDir",
                "Backup Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Error during backup:`n`n$($_.Exception.Message)",
                "Backup Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })

    $btnExit.Add_Click({
        <#
        .SYNOPSIS
            Exits the application
        .NOTES
            Closes the main menu form
        #>
        $form.Close()
    })

    #endregion

    # Show the form
    $form.ShowDialog() | Out-Null
}

#endregion

#region Main Execution

# Clear console for clean output
Clear-Host

# Display banner
Write-Host @"

╔════════════════════════════════════════════════════════════════════╗
║                                                                    ║
║           LOGISTICS AUTOMATION SUITE v2.4.1                        ║
║           JT Custom Trailers                                       ║
║                                                                    ║
║           Complete freight quote automation system                 ║
║           with analytics & performance tracking                    ║
║                                                                    ║
║           FIXED IN v2.4.1:                                         ║
║           • HTML parsing errors resolved with proper escaping      ║
║           • Enhanced error handling and user feedback              ║
║           • Improved console output and progress tracking          ║
║           • All here-strings properly formatted                    ║
║                                                                    ║
╚════════════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan

Write-Host "Initializing suite..." -ForegroundColor Green
Write-Host "Location: $script:ScriptDirectory" -ForegroundColor Gray
Write-Host ""

# Check for required and optional scripts
$requiredScripts = @(
    @{ Name = "Logistics-Config-GUI.ps1"; Required = $true; Description = "Configuration Tool" },
    @{ Name = "Freight-Recipient-Manager.ps1"; Required = $true; Description = "Recipient Manager" },
    @{ Name = "Auction-Quote-Tracker.ps1"; Required = $true; Description = "Quote Tracker" },
    @{ Name = "Generate-LogisticsEmail.ps1"; Required = $true; Description = "Email Generator" },
    @{ Name = "Convert-HTMLtoPDF.ps1"; Required = $false; Important = $true; Description = "PDF Helper" },
    @{ Name = "Logistics-Analytics-Dashboard.ps1"; Required = $false; New = $true; Description = "Analytics Dashboard" },
    @{ Name = "Generic-PDF-Invoice-Parser.ps1"; Required = $false; New = $true; Description = "PDF Invoice Parser" }
)

$missingScripts = @()
$newScripts = @()
$importantMissing = @()
$foundScripts = @()

foreach ($scriptInfo in $requiredScripts) {
    $scriptPath = Join-Path $script:ScriptDirectory $scriptInfo.Name
    $exists = Test-Path $scriptPath

    if (!$exists) {
        if ($scriptInfo.Required) {
            $missingScripts += $scriptInfo
        }
        elseif ($scriptInfo.Important) {
            $importantMissing += $scriptInfo
        }
        else {
            $newScripts += $scriptInfo
        }
    }
    else {
        $foundScripts += $scriptInfo
        if ($scriptInfo.New) {
            Write-Host "✓ Found new feature: $($scriptInfo.Description)" -ForegroundColor Green
        }
        elseif ($scriptInfo.Important) {
            Write-Host "✓ Found PDF helper: $($scriptInfo.Description)" -ForegroundColor Green
        }
    }
}

# Display warnings for missing scripts
if ($missingScripts.Count -gt 0) {
    Write-Host "`n⚠ WARNING: Required scripts are missing:" -ForegroundColor Yellow
    foreach ($missing in $missingScripts) {
        Write-Host "   ✗ $($missing.Name) - $($missing.Description)" -ForegroundColor Red
    }
    Write-Host "`nPlease ensure all required scripts are in the same directory." -ForegroundColor Yellow
    Write-Host "Press any key to continue anyway..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

if ($importantMissing.Count -gt 0) {
    Write-Host "`nIMPORTANT: PDF conversion helper is missing:" -ForegroundColor Yellow
    foreach ($missing in $importantMissing) {
        Write-Host "   ✗ $($missing.Name) - $($missing.Description)" -ForegroundColor Red
    }
    Write-Host "`nWithout this helper:" -ForegroundColor Yellow
    Write-Host "   • HTML reports will be created successfully" -ForegroundColor Gray
    Write-Host "   • You'll need to manually print HTML to PDF in browser (Ctrl+P)" -ForegroundColor Gray
    Write-Host "   • This is actually fast and reliable!" -ForegroundColor Green
    Write-Host ""
}

if ($newScripts.Count -gt 0) {
    Write-Host "`nTIP: Optional features are available but not installed:" -ForegroundColor Cyan
    foreach ($new in $newScripts) {
        Write-Host "   - $($new.Name) - $($new.Description)" -ForegroundColor Gray
    }
    Write-Host "   Add these scripts to unlock additional functionality!" -ForegroundColor Yellow
    Write-Host ""
}

# Launch main menu
Write-Host "Starting main menu..." -ForegroundColor Green
Write-Host ""

Show-EnhancedMainMenu

#endregion