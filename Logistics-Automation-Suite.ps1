<#
.SYNOPSIS
    Logistics Automation Suite - Enhanced Master Launcher v2.0
    
.DESCRIPTION
    Updated central hub with new analytics dashboard, enhanced carrier tracking,
    and universal PDF invoice parser integration.
    
.EXAMPLE
    .\Update-Logistics-Suite.ps1
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 2025-01-08
    Version: 2.0.0
    Change Date: 2025-01-08
    Change Purpose: Added Analytics Dashboard & PDF Parser Integration
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ScriptDirectory = $PSScriptRoot

#region Main Menu GUI
function Show-EnhancedMainMenu {
    <#
    .SYNOPSIS
        Shows the enhanced main menu with new features
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-08
        Version: 2.0.0
        Change Date: 2025-01-08
        Change Purpose: Added new tool buttons
    #>
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Logistics Automation Suite v2.0 - JT Custom Trailers"
    $form.Size = New-Object System.Drawing.Size(950, 750)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::White
    
    # Header
    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Location = New-Object System.Drawing.Point(0, 0)
    $pnlHeader.Size = New-Object System.Drawing.Size(950, 100)
    $pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(44, 62, 80)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(900, 35)
    $lblTitle.Text = "Logistics Automation Suite v2.0"
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
    $lblVersion.Text = "Analytics Dashboard • Enhanced PDF Parser • Carrier Performance Metrics"
    $lblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblVersion.ForeColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $pnlHeader.Controls.Add($lblVersion)
    
    $form.Controls.Add($pnlHeader)
    
    # Main Tools Section
    $yPos = 120
    
    $lblMainTools = New-Object System.Windows.Forms.Label
    $lblMainTools.Location = New-Object System.Drawing.Point(30, $yPos)
    $lblMainTools.Size = New-Object System.Drawing.Size(400, 25)
    $lblMainTools.Text = "Main Tools"
    $lblMainTools.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblMainTools)
    
    $yPos += 40
    
    # Row 1
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
    
    # Row 2
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
    
    # Quick Actions Section
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
    
    # Utilities Section
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
    
    # Help & Documentation
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
4. Click "Generate Email" to create quote requests
5. Track responses and analyze costs in "Quote Tracker" and "Analytics Dashboard"
"@
    $lblHelp.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblHelp.ForeColor = [System.Drawing.Color]::FromArgb(52, 73, 94)
    $grpHelp.Controls.Add($lblHelp)
    
    $form.Controls.Add($grpHelp)
    
    # Footer Buttons
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Location = New-Object System.Drawing.Point(790, 670)
    $btnExit.Size = New-Object System.Drawing.Size(120, 35)
    $btnExit.Text = "Exit"
    $btnExit.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($btnExit)
    
    #region Event Handlers
    
    $btnConfig.Add_Click({
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
        # Prompt for CSV and config files
        $openCSV = New-Object System.Windows.Forms.OpenFileDialog
        $openCSV.Filter = "CSV files (*.csv)|*.csv"
        $openCSV.Title = "Select Auction Lots CSV File"
        
        if ($openCSV.ShowDialog() -eq "OK") {
            $openConfig = New-Object System.Windows.Forms.OpenFileDialog
            $openConfig.Filter = "JSON files (*.json)|*.json"
            $openConfig.Title = "Select Configuration File"
            $openConfig.InitialDirectory = Join-Path $script:ScriptDirectory "Templates"
            
            if ($openConfig.ShowDialog() -eq "OK") {
                $imageDir = New-Object System.Windows.Forms.FolderBrowserDialog
                $imageDir.Description = "Select Image Directory"
                
                if ($imageDir.ShowDialog() -eq "OK") {
                    $genScript = Join-Path $script:ScriptDirectory "Generate-LogisticsEmail.ps1"
                    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$genScript`" -CSVPath `"$($openCSV.FileName)`" -ImageDirectory `"$($imageDir.SelectedPath)`" -CreateOutlookDraft"
                    
                    Start-Process powershell.exe -ArgumentList $arguments
                }
            }
        }
    })
    
    $btnTracker.Add_Click({
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
        $analyticsScript = Join-Path $script:ScriptDirectory "Logistics-Analytics-Dashboard.ps1"
        if (Test-Path $analyticsScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$analyticsScript`""
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Analytics Dashboard script not found:`n$analyticsScript`n`nPlease ensure the new analytics script is in the suite directory.",
                "Script Not Found",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    })
    
    $btnQuickBrolyn.Add_Click({
        $configScript = Join-Path $script:ScriptDirectory "Logistics-Config-GUI.ps1"
        if (Test-Path $configScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$configScript`" -LoadTemplate `"Brolyn_Auctions`""
        }
    })
    
    $btnQuickGeneric.Add_Click({
        $configScript = Join-Path $script:ScriptDirectory "Logistics-Config-GUI.ps1"
        if (Test-Path $configScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$configScript`" -LoadTemplate `"Generic_Template`""
        }
    })
    
    $btnParseInvoice.Add_Click({
        $openPDF = New-Object System.Windows.Forms.OpenFileDialog
        $openPDF.Filter = "PDF files (*.pdf)|*.pdf"
        $openPDF.Title = "Select Invoice PDF to Parse"
        
        if ($openPDF.ShowDialog() -eq "OK") {
            $parserScript = Join-Path $script:ScriptDirectory "Generic-PDF-Invoice-Parser.ps1"
            if (Test-Path $parserScript) {
                Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$parserScript`" -PDFPath `"$($openPDF.FileName)`" -OutputFormat Config"
            }
        }
    })
    
    $btnTemplates.Add_Click({
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
        $analyticsScript = Join-Path $script:ScriptDirectory "Logistics-Analytics-Dashboard.ps1"
        if (Test-Path $analyticsScript) {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$analyticsScript`""
        }
    })
    
    $btnDataBackup.Add_Click({
        $backupDir = ".\Backups\Backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        
        if (!(Test-Path ".\Backups")) {
            New-Item -ItemType Directory -Path ".\Backups" -Force | Out-Null
        }
        
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        
        # Backup data files
        $dataFiles = @(
            ".\Data\AuctionQuotes.json",
            ".\Data\FreightRecipients.json",
            ".\Data\InvoicePatterns.json"
        )
        
        $backedUp = 0
        foreach ($file in $dataFiles) {
            if (Test-Path $file) {
                Copy-Item $file -Destination $backupDir -Force
                $backedUp++
            }
        }
        
        # Backup templates
        if (Test-Path ".\Templates") {
            Copy-Item ".\Templates\*" -Destination (Join-Path $backupDir "Templates") -Force -Recurse
        }
        
        [System.Windows.Forms.MessageBox]::Show(
            "Backup complete!`n`nBacked up $backedUp data files`n`nLocation: $backupDir",
            "Backup Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    
    $btnExit.Add_Click({
        $form.Close()
    })
    
    #endregion
    
    $form.ShowDialog() | Out-Null
}
#endregion

#region Main Execution
Clear-Host

Write-Host @"

╔════════════════════════════════════════════════════════════════════╗
║                                                                    ║
║           LOGISTICS AUTOMATION SUITE v2.0                          ║
║           JT Custom Trailers                                       ║
║                                                                    ║
║           Complete freight quote automation system                 ║
║           with analytics & performance tracking                    ║
║                                                                    ║
║           NEW IN v2.0:                                             ║
║           • Analytics Dashboard with cost trends                   ║
║           • Carrier performance metrics & tracking                 ║
║           • Universal PDF invoice parser                           ║
║                                                                    ║
╚════════════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan

Write-Host "Initializing suite..." -ForegroundColor Green
Write-Host ""

# Check for required scripts
$requiredScripts = @(
    @{ Name = "Logistics-Config-GUI.ps1"; Required = $true },
    @{ Name = "Freight-Recipient-Manager.ps1"; Required = $true },
    @{ Name = "Auction-Quote-Tracker.ps1"; Required = $true },
    @{ Name = "Generate-LogisticsEmail.ps1"; Required = $true },
    @{ Name = "Logistics-Analytics-Dashboard.ps1"; Required = $false; New = $true },
    @{ Name = "Generic-PDF-Invoice-Parser.ps1"; Required = $false; New = $true }
)

$missingScripts = @()
$newScripts = @()

foreach ($scriptInfo in $requiredScripts) {
    $scriptPath = Join-Path $script:ScriptDirectory $scriptInfo.Name
    if (!(Test-Path $scriptPath)) {
        if ($scriptInfo.Required) {
            $missingScripts += $scriptInfo.Name
        }
        else {
            $newScripts += $scriptInfo.Name
        }
    }
    elseif ($scriptInfo.New) {
        Write-Host "Found new feature: $($scriptInfo.Name)" -ForegroundColor Green
    }
}

if ($missingScripts.Count -gt 0) {
    Write-Host "WARNING: Required scripts are missing:" -ForegroundColor Yellow
    foreach ($missing in $missingScripts) {
        Write-Host "   - $missing" -ForegroundColor Red
    }
    Write-Host "`nPlease ensure all scripts are in the same directory." -ForegroundColor Yellow
    Write-Host "Press any key to continue anyway..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

if ($newScripts.Count -gt 0) {
    Write-Host "`nTIP: New optional features are available but not installed:" -ForegroundColor Cyan
    foreach ($new in $newScripts) {
        Write-Host "   - $new" -ForegroundColor Gray
    }
    Write-Host "   Add these scripts to unlock analytics and enhanced PDF parsing!" -ForegroundColor Yellow
    Write-Host ""
}

Show-EnhancedMainMenu
#endregion
