<#
.SYNOPSIS
    Logistics Automation Suite - Master Launcher
    
.DESCRIPTION
    Central hub for all logistics automation tools. Provides easy access to:
    - Configuration GUI
    - Recipient Manager
    - Quote Tracker
    - Email Generation
    - Quick actions and workflows
    
.EXAMPLE
    .\Logistics-Automation-Suite.ps1
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 2025-01-07
    Version: 1.0.0
    Change Date: 
    Change Purpose: Initial Release - Master Suite Integration
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ScriptDirectory = $PSScriptRoot

#region Main Menu GUI
function Show-MainMenu {
    <#
    .SYNOPSIS
        Shows the main menu for the logistics automation suite
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Logistics Automation Suite - JT Custom Trailers"
    $form.Size = New-Object System.Drawing.Size(900, 650)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::White
    
    # Header
    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Location = New-Object System.Drawing.Point(0, 0)
    $pnlHeader.Size = New-Object System.Drawing.Size(900, 80)
    $pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(44, 62, 80)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(850, 30)
    $lblTitle.Text = "🚛 Logistics Automation Suite"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $pnlHeader.Controls.Add($lblTitle)
    
    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(20, 50)
    $lblSubtitle.Size = New-Object System.Drawing.Size(850, 20)
    $lblSubtitle.Text = "Complete freight quote automation for JT Custom Trailers"
    $lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblSubtitle.ForeColor = [System.Drawing.Color]::LightGray
    $pnlHeader.Controls.Add($lblSubtitle)
    
    $form.Controls.Add($pnlHeader)
    
    # Main Tools Section
    $yPos = 100
    
    $lblMainTools = New-Object System.Windows.Forms.Label
    $lblMainTools.Location = New-Object System.Drawing.Point(30, $yPos)
    $lblMainTools.Size = New-Object System.Drawing.Size(400, 25)
    $lblMainTools.Text = "📌 Main Tools"
    $lblMainTools.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblMainTools)
    
    $yPos += 40
    
    # Configuration GUI Button
    $btnConfig = New-Object System.Windows.Forms.Button
    $btnConfig.Location = New-Object System.Drawing.Point(30, $yPos)
    $btnConfig.Size = New-Object System.Drawing.Size(400, 60)
    $btnConfig.Text = "⚙️ Configuration Tool`n   Create auction configs & import from PDF invoices"
    $btnConfig.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $btnConfig.ForeColor = [System.Drawing.Color]::White
    $btnConfig.FlatStyle = "Flat"
    $btnConfig.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnConfig.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($btnConfig)
    
    # Recipient Manager Button
    $btnRecipients = New-Object System.Windows.Forms.Button
    $btnRecipients.Location = New-Object System.Drawing.Point(450, $yPos)
    $btnRecipients.Size = New-Object System.Drawing.Size(400, 60)
    $btnRecipients.Text = "📧 Recipient Manager`n   Manage freight company contacts & favorites"
    $btnRecipients.BackColor = [System.Drawing.Color]::FromArgb(46, 204, 113)
    $btnRecipients.ForeColor = [System.Drawing.Color]::White
    $btnRecipients.FlatStyle = "Flat"
    $btnRecipients.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnRecipients.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($btnRecipients)
    
    $yPos += 80
    
    # Email Generator Button
    $btnGenerate = New-Object System.Windows.Forms.Button
    $btnGenerate.Location = New-Object System.Drawing.Point(30, $yPos)
    $btnGenerate.Size = New-Object System.Drawing.Size(400, 60)
    $btnGenerate.Text = "📨 Generate Email`n   Create logistics quote email with attachments"
    $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(155, 89, 182)
    $btnGenerate.ForeColor = [System.Drawing.Color]::White
    $btnGenerate.FlatStyle = "Flat"
    $btnGenerate.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnGenerate.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($btnGenerate)
    
    # Quote Tracker Button
    $btnTracker = New-Object System.Windows.Forms.Button
    $btnTracker.Location = New-Object System.Drawing.Point(450, $yPos)
    $btnTracker.Size = New-Object System.Drawing.Size(400, 60)
    $btnTracker.Text = "📊 Quote Tracker`n   Track auctions, quotes, and freight costs"
    $btnTracker.BackColor = [System.Drawing.Color]::FromArgb(230, 126, 34)
    $btnTracker.ForeColor = [System.Drawing.Color]::White
    $btnTracker.FlatStyle = "Flat"
    $btnTracker.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnTracker.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $form.Controls.Add($btnTracker)
    
    $yPos += 100
    
    # Quick Actions Section
    $lblQuickActions = New-Object System.Windows.Forms.Label
    $lblQuickActions.Location = New-Object System.Drawing.Point(30, $yPos)
    $lblQuickActions.Size = New-Object System.Drawing.Size(400, 25)
    $lblQuickActions.Text = "⚡ Quick Actions"
    $lblQuickActions.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($lblQuickActions)
    
    $yPos += 40
    
    # Quick Workflow Buttons
    $btnQuickBrolyn = New-Object System.Windows.Forms.Button
    $btnQuickBrolyn.Location = New-Object System.Drawing.Point(30, $yPos)
    $btnQuickBrolyn.Size = New-Object System.Drawing.Size(260, 45)
    $btnQuickBrolyn.Text = "🚀 New Brolyn Auction"
    $btnQuickBrolyn.BackColor = [System.Drawing.Color]::LightGreen
    $btnQuickBrolyn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnQuickBrolyn)
    
    $btnQuickGeneric = New-Object System.Windows.Forms.Button
    $btnQuickGeneric.Location = New-Object System.Drawing.Point(300, $yPos)
    $btnQuickGeneric.Size = New-Object System.Drawing.Size(260, 45)
    $btnQuickGeneric.Text = "🆕 New Generic Auction"
    $btnQuickGeneric.BackColor = [System.Drawing.Color]::LightBlue
    $btnQuickGeneric.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnQuickGeneric)
    
    $btnTemplates = New-Object System.Windows.Forms.Button
    $btnTemplates.Location = New-Object System.Drawing.Point(570, $yPos)
    $btnTemplates.Size = New-Object System.Drawing.Size(280, 45)
    $btnTemplates.Text = "📋 Create Starter Templates"
    $btnTemplates.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.Controls.Add($btnTemplates)
    
    $yPos += 60
    
    # Help & Documentation
    $lblHelp = New-Object System.Windows.Forms.Label
    $lblHelp.Location = New-Object System.Drawing.Point(30, $yPos)
    $lblHelp.Size = New-Object System.Drawing.Size(820, 60)
    $lblHelp.Text = @"
💡 Getting Started:
1. Click "Configuration Tool" to set up auction details (or import from PDF invoice)
2. Use "Recipient Manager" to add freight companies
3. Click "Generate Email" to create your quote request
4. Track everything in "Quote Tracker" once quotes come back
"@
    $lblHelp.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblHelp.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($lblHelp)
    
    # Footer Buttons
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Location = New-Object System.Drawing.Point(750, 560)
    $btnExit.Size = New-Object System.Drawing.Size(120, 35)
    $btnExit.Text = "✖️ Exit"
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
                "Configuration script not found:`n$configScript`n`nPlease ensure all scripts are in the same directory.",
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
                $imageDir = [System.Windows.Forms.FolderBrowserDialog]::new()
                $imageDir.Description = "Select Image Directory"
                
                if ($imageDir.ShowDialog() -eq "OK") {
                    $genScript = Join-Path $script:ScriptDirectory "Integrated-LogisticsEmail.ps1"
                    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$genScript`" -CSVPath `"$($openCSV.FileName)`" -ConfigPath `"$($openConfig.FileName)`" -ImageDirectory `"$($imageDir.SelectedPath)`" -CreateOutlookDraft"
                    
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
║           LOGISTICS AUTOMATION SUITE v1.0                          ║
║           JT Custom Trailers                                       ║
║                                                                    ║
║           Complete freight quote automation system                 ║
║                                                                    ║
╚════════════════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan

Write-Host "Initializing suite..." -ForegroundColor Green
Write-Host ""

# Check for required scripts
$requiredScripts = @(
    "Logistics-Config-GUI.ps1",
    "Freight-Recipient-Manager.ps1",
    "Auction-Quote-Tracker.ps1",
    "Integrated-LogisticsEmail.ps1"
)

$missingScripts = @()
foreach ($script in $requiredScripts) {
    $scriptPath = Join-Path $script:ScriptDirectory $script
    if (!(Test-Path $scriptPath)) {
        $missingScripts += $script
    }
}

if ($missingScripts.Count -gt 0) {
    Write-Host "⚠️  WARNING: Some scripts are missing:" -ForegroundColor Yellow
    foreach ($missing in $missingScripts) {
        Write-Host "   - $missing" -ForegroundColor Red
    }
    Write-Host "`nPlease ensure all scripts are in the same directory." -ForegroundColor Yellow
    Write-Host "Press any key to continue anyway..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Show-MainMenu
#endregion
