<#
.SYNOPSIS
    Auction Quote Tracker
    
.DESCRIPTION
    Comprehensive tracking system for auction logistics quotes. Tracks quotes sent,
    responses received, costs, selected carriers, and overall status. Provides
    analytics and reporting on freight costs and carrier performance.
    
.EXAMPLE
    .\Auction-Quote-Tracker.ps1
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 2025-01-07
    Version: 1.0.0
    Change Date: 
    Change Purpose: Initial Release - Quote Tracking System
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:AuctionsFile = ".\Data\AuctionQuotes.json"
$script:Auctions = @()

#region Data Management Functions
function Initialize-TrackerData {
    <#
    .SYNOPSIS
        Initializes auction tracker data storage
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    $dataDir = Split-Path $script:AuctionsFile -Parent
    if (!(Test-Path $dataDir)) {
        New-Item -ItemType Directory -Path $dataDir -Force | Out-Null
    }
    
    if (Test-Path $script:AuctionsFile) {
        $script:Auctions = Get-Content $script:AuctionsFile -Raw | ConvertFrom-Json
        Write-Host "✅ Loaded $($script:Auctions.Count) auction records" -ForegroundColor Green
    }
    else {
        $script:Auctions = @()
        Save-TrackerData
    }
}

function Save-TrackerData {
    <#
    .SYNOPSIS
        Saves tracker data to JSON file
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    $script:Auctions | ConvertTo-Json -Depth 10 | Out-File $script:AuctionsFile -Encoding UTF8
}

function Add-AuctionRecord {
    <#
    .SYNOPSIS
        Adds a new auction tracking record
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [hashtable]$AuctionData
    )
    
    $AuctionData.Id = (New-Guid).ToString()
    $AuctionData.CreatedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $AuctionData.ModifiedDate = $AuctionData.CreatedDate
    
    if (!$AuctionData.Quotes) { $AuctionData.Quotes = @() }
    if (!$AuctionData.Status) { $AuctionData.Status = "Pending Quotes" }
    
    $script:Auctions += $AuctionData
    Save-TrackerData
    
    return $AuctionData.Id
}

function Update-AuctionRecord {
    <#
    .SYNOPSIS
        Updates existing auction record
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$Id,
        [hashtable]$UpdatedData
    )
    
    for ($i = 0; $i -lt $script:Auctions.Count; $i++) {
        if ($script:Auctions[$i].Id -eq $Id) {
            foreach ($key in $UpdatedData.Keys) {
                $script:Auctions[$i].$key = $UpdatedData[$key]
            }
            $script:Auctions[$i].ModifiedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Save-TrackerData
            return $true
        }
    }
    return $false
}

function Add-QuoteToAuction {
    <#
    .SYNOPSIS
        Adds a quote response to an auction record
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$AuctionId,
        [hashtable]$QuoteData
    )
    
    for ($i = 0; $i -lt $script:Auctions.Count; $i++) {
        if ($script:Auctions[$i].Id -eq $AuctionId) {
            $QuoteData.Id = (New-Guid).ToString()
            $QuoteData.ReceivedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            
            if (!$script:Auctions[$i].Quotes) {
                $script:Auctions[$i].Quotes = @()
            }
            
            $script:Auctions[$i].Quotes += $QuoteData
            $script:Auctions[$i].ModifiedDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            
            # Auto-update status
            if ($script:Auctions[$i].Status -eq "Pending Quotes") {
                $script:Auctions[$i].Status = "Quotes Received"
            }
            
            Save-TrackerData
            return $true
        }
    }
    return $false
}

function Get-AuctionStatistics {
    <#
    .SYNOPSIS
        Gets statistics about auctions and quotes
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    $stats = @{
        TotalAuctions = $script:Auctions.Count
        PendingQuotes = ($script:Auctions | Where-Object { $_.Status -eq "Pending Quotes" }).Count
        QuotesReceived = ($script:Auctions | Where-Object { $_.Status -eq "Quotes Received" }).Count
        Booked = ($script:Auctions | Where-Object { $_.Status -eq "Booked" }).Count
        Completed = ($script:Auctions | Where-Object { $_.Status -eq "Completed" }).Count
        TotalFreightCost = 0
        AverageQuoteTime = 0
        TopCarriers = @()
    }
    
    # Calculate total freight cost
    foreach ($auction in $script:Auctions) {
        if ($auction.SelectedQuote -and $auction.SelectedQuote.Amount) {
            $stats.TotalFreightCost += $auction.SelectedQuote.Amount
        }
    }
    
    return $stats
}

function Export-AuctionReport {
    <#
    .SYNOPSIS
        Exports auction data to CSV
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param (
        [string]$OutputPath
    )
    
    $exportData = $script:Auctions | ForEach-Object {
        $quoteCount = if ($_.Quotes) { $_.Quotes.Count } else { 0 }
        $selectedCarrier = if ($_.SelectedQuote) { $_.SelectedQuote.Company } else { "None" }
        $selectedAmount = if ($_.SelectedQuote) { $_.SelectedQuote.Amount } else { 0 }
        
        [PSCustomObject]@{
            AuctionDate = $_.AuctionDate
            AuctionCompany = $_.AuctionCompany
            PickupLocation = $_.PickupLocation
            TotalLots = $_.TotalLots
            TotalPallets = $_.TotalPallets
            Status = $_.Status
            QuotesReceived = $quoteCount
            SelectedCarrier = $selectedCarrier
            FreightCost = $selectedAmount
            Notes = $_.Notes
            CreatedDate = $_.CreatedDate
        }
    }
    
    $exportData | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Host "✅ Exported to: $OutputPath" -ForegroundColor Green
}
#endregion

#region GUI Functions
function New-TrackerGUI {
    <#
    .SYNOPSIS
        Creates the main tracker GUI
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-07
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Auction Quote Tracker"
    $form.Size = New-Object System.Drawing.Size(1400, 750)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    #region Top Panel - Stats
    $pnlStats = New-Object System.Windows.Forms.Panel
    $pnlStats.Location = New-Object System.Drawing.Point(10, 10)
    $pnlStats.Size = New-Object System.Drawing.Size(1365, 80)
    $pnlStats.BorderStyle = "FixedSingle"
    $pnlStats.BackColor = [System.Drawing.Color]::WhiteSmoke
    
    $lblStatsTitle = New-Object System.Windows.Forms.Label
    $lblStatsTitle.Location = New-Object System.Drawing.Point(10, 10)
    $lblStatsTitle.Size = New-Object System.Drawing.Size(200, 20)
    $lblStatsTitle.Text = "📊 Quick Statistics"
    $lblStatsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $pnlStats.Controls.Add($lblStatsTitle)
    
    $lblStats = New-Object System.Windows.Forms.Label
    $lblStats.Location = New-Object System.Drawing.Point(10, 35)
    $lblStats.Size = New-Object System.Drawing.Size(1340, 40)
    $lblStats.Text = "Loading statistics..."
    $pnlStats.Controls.Add($lblStats)
    
    $form.Controls.Add($pnlStats)
    #endregion
    
    #region Filter Controls
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(20, 105)
    $lblStatus.Size = New-Object System.Drawing.Size(60, 20)
    $lblStatus.Text = "Status:"
    $form.Controls.Add($lblStatus)
    
    $cmbStatus = New-Object System.Windows.Forms.ComboBox
    $cmbStatus.Location = New-Object System.Drawing.Point(85, 103)
    $cmbStatus.Size = New-Object System.Drawing.Size(180, 25)
    $cmbStatus.DropDownStyle = "DropDownList"
    @("All", "Pending Quotes", "Quotes Received", "Booked", "In Transit", "Completed", "Cancelled") | ForEach-Object {
        $cmbStatus.Items.Add($_) | Out-Null
    }
    $cmbStatus.SelectedIndex = 0
    $form.Controls.Add($cmbStatus)
    
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(280, 105)
    $lblSearch.Size = New-Object System.Drawing.Size(60, 20)
    $lblSearch.Text = "Search:"
    $form.Controls.Add($lblSearch)
    
    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(345, 103)
    $txtSearch.Size = New-Object System.Drawing.Size(250, 25)
    $form.Controls.Add($txtSearch)
    
    $btnSearch = New-Object System.Windows.Forms.Button
    $btnSearch.Location = New-Object System.Drawing.Point(600, 102)
    $btnSearch.Size = New-Object System.Drawing.Size(80, 27)
    $btnSearch.Text = "🔍 Search"
    $form.Controls.Add($btnSearch)
    
    $btnAddNew = New-Object System.Windows.Forms.Button
    $btnAddNew.Location = New-Object System.Drawing.Point(1100, 100)
    $btnAddNew.Size = New-Object System.Drawing.Size(140, 32)
    $btnAddNew.Text = "➕ New Auction"
    $btnAddNew.BackColor = [System.Drawing.Color]::LightGreen
    $btnAddNew.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($btnAddNew)
    
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Location = New-Object System.Drawing.Point(1250, 100)
    $btnRefresh.Size = New-Object System.Drawing.Size(125, 32)
    $btnRefresh.Text = "🔄 Refresh"
    $form.Controls.Add($btnRefresh)
    #endregion
    
    #region DataGridView
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location = New-Object System.Drawing.Point(20, 145)
    $dgv.Size = New-Object System.Drawing.Size(1355, 480)
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.ReadOnly = $true
    $dgv.SelectionMode = "FullRowSelect"
    $dgv.MultiSelect = $false
    $dgv.AutoSizeColumnsMode = "Fill"
    
    # Define columns
    $dgv.Columns.Add("AuctionDate", "Date") | Out-Null
    $dgv.Columns.Add("AuctionCompany", "Auction Co.") | Out-Null
    $dgv.Columns.Add("PickupLocation", "Pickup Location") | Out-Null
    $dgv.Columns.Add("TotalLots", "Lots") | Out-Null
    $dgv.Columns.Add("Pallets", "Pallets") | Out-Null
    $dgv.Columns.Add("Status", "Status") | Out-Null
    $dgv.Columns.Add("QuotesReceived", "Quotes") | Out-Null
    $dgv.Columns.Add("SelectedCarrier", "Carrier") | Out-Null
    $dgv.Columns.Add("FreightCost", "Cost") | Out-Null
    
    $dgv.Columns["AuctionDate"].Width = 100
    $dgv.Columns["TotalLots"].Width = 60
    $dgv.Columns["Pallets"].Width = 70
    $dgv.Columns["Status"].Width = 120
    $dgv.Columns["QuotesReceived"].Width = 70
    $dgv.Columns["FreightCost"].Width = 90
    
    $form.Controls.Add($dgv)
    #endregion
    
    #region Bottom Buttons
    $btnView = New-Object System.Windows.Forms.Button
    $btnView.Location = New-Object System.Drawing.Point(20, 640)
    $btnView.Size = New-Object System.Drawing.Size(140, 35)
    $btnView.Text = "👁️ View Details"
    $btnView.BackColor = [System.Drawing.Color]::LightBlue
    $form.Controls.Add($btnView)
    
    $btnAddQuote = New-Object System.Windows.Forms.Button
    $btnAddQuote.Location = New-Object System.Drawing.Point(170, 640)
    $btnAddQuote.Size = New-Object System.Drawing.Size(140, 35)
    $btnAddQuote.Text = "💵 Add Quote"
    $btnAddQuote.BackColor = [System.Drawing.Color]::LightGreen
    $form.Controls.Add($btnAddQuote)
    
    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Location = New-Object System.Drawing.Point(320, 640)
    $btnEdit.Size = New-Object System.Drawing.Size(120, 35)
    $btnEdit.Text = "✏️ Edit"
    $form.Controls.Add($btnEdit)
    
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Location = New-Object System.Drawing.Point(450, 640)
    $btnDelete.Size = New-Object System.Drawing.Size(120, 35)
    $btnDelete.Text = "🗑️ Delete"
    $btnDelete.BackColor = [System.Drawing.Color]::LightCoral
    $form.Controls.Add($btnDelete)
    
    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Location = New-Object System.Drawing.Point(1100, 640)
    $btnExport.Size = New-Object System.Drawing.Size(140, 35)
    $btnExport.Text = "📤 Export Report"
    $form.Controls.Add($btnExport)
    
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(1250, 640)
    $btnClose.Size = New-Object System.Drawing.Size(125, 35)
    $btnClose.Text = "✖️ Close"
    $form.Controls.Add($btnClose)
    #endregion
    
    #region Helper Functions
    function Update-Statistics {
        $stats = Get-AuctionStatistics
        $lblStats.Text = "Total Auctions: $($stats.TotalAuctions)  |  Pending: $($stats.PendingQuotes)  |  Quotes Received: $($stats.QuotesReceived)  |  Booked: $($stats.Booked)  |  Completed: $($stats.Completed)  |  Total Freight Cost: `$$($stats.TotalFreightCost.ToString('N2'))"
    }
    
    function Refresh-AuctionGrid {
        param (
            [string]$StatusFilter = "All",
            [string]$SearchText = ""
        )
        
        $dgv.Rows.Clear()
        $filtered = $script:Auctions
        
        if ($StatusFilter -ne "All") {
            $filtered = $filtered | Where-Object { $_.Status -eq $StatusFilter }
        }
        
        if ($SearchText) {
            $filtered = $filtered | Where-Object {
                $_.AuctionCompany -like "*$SearchText*" -or
                $_.PickupLocation -like "*$SearchText*" -or
                $_.Notes -like "*$SearchText*"
            }
        }
        
        # Sort by date descending
        $filtered = $filtered | Sort-Object AuctionDate -Descending
        
        foreach ($auction in $filtered) {
            $quoteCount = if ($auction.Quotes) { $auction.Quotes.Count } else { 0 }
            $selectedCarrier = if ($auction.SelectedQuote) { $auction.SelectedQuote.Company } else { "-" }
            $freightCost = if ($auction.SelectedQuote -and $auction.SelectedQuote.Amount) {
                "`$$($auction.SelectedQuote.Amount.ToString('N2'))"
            } else { "-" }
            
            $row = @(
                $auction.AuctionDate,
                $auction.AuctionCompany,
                $auction.PickupLocation,
                $auction.TotalLots,
                $auction.TotalPallets,
                $auction.Status,
                $quoteCount,
                $selectedCarrier,
                $freightCost
            )
            $dgv.Rows.Add($row) | Out-Null
            $dgv.Rows[$dgv.Rows.Count - 1].Tag = $auction.Id
            
            # Color code by status
            $rowIndex = $dgv.Rows.Count - 1
            switch ($auction.Status) {
                "Pending Quotes" { $dgv.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow }
                "Quotes Received" { $dgv.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightCyan }
                "Booked" { $dgv.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen }
                "Completed" { $dgv.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray }
            }
        }
        
        Update-Statistics
    }
    
    function Show-AuctionDetailsDialog {
        param (
            [string]$AuctionId
        )
        
        $auction = $script:Auctions | Where-Object { $_.Id -eq $AuctionId }
        if (!$auction) { return }
        
        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = "Auction Details - $($auction.AuctionCompany)"
        $dlg.Size = New-Object System.Drawing.Size(800, 700)
        $dlg.StartPosition = "CenterParent"
        $dlg.FormBorderStyle = "Sizable"
        
        # Create TabControl
        $tabControl = New-Object System.Windows.Forms.TabControl
        $tabControl.Location = New-Object System.Drawing.Point(10, 10)
        $tabControl.Size = New-Object System.Drawing.Size(760, 590)
        
        # Tab 1: Basic Info
        $tabInfo = New-Object System.Windows.Forms.TabPage
        $tabInfo.Text = "Auction Info"
        
        $txtInfo = New-Object System.Windows.Forms.TextBox
        $txtInfo.Location = New-Object System.Drawing.Point(10, 10)
        $txtInfo.Size = New-Object System.Drawing.Size(730, 540)
        $txtInfo.Multiline = $true
        $txtInfo.ScrollBars = "Vertical"
        $txtInfo.ReadOnly = $true
        $txtInfo.Font = New-Object System.Drawing.Font("Consolas", 9)
        
        $infoText = @"
AUCTION INFORMATION

Date: $($auction.AuctionDate)
Auction Company: $($auction.AuctionCompany)
Pickup Location: $($auction.PickupLocation)
Delivery Location: $($auction.DeliveryLocation)

SHIPMENT DETAILS
Total Lots: $($auction.TotalLots)
Total Pallets: $($auction.TotalPallets)
Truck Types: $($auction.TruckTypes)

STATUS
Current Status: $($auction.Status)
Created: $($auction.CreatedDate)
Last Modified: $($auction.ModifiedDate)

NOTES
$($auction.Notes)

CONFIGURATION FILE
$($auction.ConfigFile)
"@
        
        $txtInfo.Text = $infoText
        $tabInfo.Controls.Add($txtInfo)
        $tabControl.TabPages.Add($tabInfo)
        
        # Tab 2: Quotes
        $tabQuotes = New-Object System.Windows.Forms.TabPage
        $tabQuotes.Text = "Quotes ($($auction.Quotes.Count))"
        
        $lstQuotes = New-Object System.Windows.Forms.ListBox
        $lstQuotes.Location = New-Object System.Drawing.Point(10, 10)
        $lstQuotes.Size = New-Object System.Drawing.Size(300, 540)
        $lstQuotes.Font = New-Object System.Drawing.Font("Consolas", 9)
        
        foreach ($quote in $auction.Quotes) {
            $quoteDisplay = "$($quote.Company) - `$$($quote.Amount) - $($quote.ReceivedDate)"
            $lstQuotes.Items.Add($quoteDisplay) | Out-Null
        }
        
        $tabQuotes.Controls.Add($lstQuotes)
        
        $txtQuoteDetail = New-Object System.Windows.Forms.TextBox
        $txtQuoteDetail.Location = New-Object System.Drawing.Point(320, 10)
        $txtQuoteDetail.Size = New-Object System.Drawing.Size(420, 540)
        $txtQuoteDetail.Multiline = $true
        $txtQuoteDetail.ScrollBars = "Vertical"
        $txtQuoteDetail.ReadOnly = $true
        $txtQuoteDetail.Font = New-Object System.Drawing.Font("Consolas", 9)
        $tabQuotes.Controls.Add($txtQuoteDetail)
        
        $lstQuotes.Add_SelectedIndexChanged({
            if ($lstQuotes.SelectedIndex -ge 0) {
                $selectedQuote = $auction.Quotes[$lstQuotes.SelectedIndex]
                $quoteDetail = @"
QUOTE DETAILS

Company: $($selectedQuote.Company)
Contact: $($selectedQuote.ContactName)
Email: $($selectedQuote.Email)
Phone: $($selectedQuote.Phone)

PRICING
Amount: `$$($selectedQuote.Amount)
Currency: USD
Received: $($selectedQuote.ReceivedDate)

DETAILS
$($selectedQuote.Notes)
"@
                $txtQuoteDetail.Text = $quoteDetail
            }
        })
        
        $tabControl.TabPages.Add($tabQuotes)
        
        $dlg.Controls.Add($tabControl)
        
        # Close button
        $btnCloseDetail = New-Object System.Windows.Forms.Button
        $btnCloseDetail.Location = New-Object System.Drawing.Point(660, 610)
        $btnCloseDetail.Size = New-Object System.Drawing.Size(110, 30)
        $btnCloseDetail.Text = "Close"
        $btnCloseDetail.DialogResult = "OK"
        $dlg.Controls.Add($btnCloseDetail)
        
        $dlg.ShowDialog() | Out-Null
    }
    #endregion
    
    #region Event Handlers
    
    # Initial load
    Refresh-AuctionGrid
    
    # Filter changes
    $cmbStatus.Add_SelectedIndexChanged({
        Refresh-AuctionGrid -StatusFilter $cmbStatus.Text -SearchText $txtSearch.Text
    })
    
    $btnSearch.Add_Click({
        Refresh-AuctionGrid -StatusFilter $cmbStatus.Text -SearchText $txtSearch.Text
    })
    
    $btnRefresh.Add_Click({
        Initialize-TrackerData
        Refresh-AuctionGrid -StatusFilter $cmbStatus.Text -SearchText $txtSearch.Text
    })
    
    # View Details
    $btnView.Add_Click({
        if ($dgv.SelectedRows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please select an auction to view.",
                "No Selection",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
        $selectedId = $dgv.SelectedRows[0].Tag
        Show-AuctionDetailsDialog -AuctionId $selectedId
    })
    
    # Export
    $btnExport.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv"
        $saveDialog.FileName = "AuctionQuotes_$(Get-Date -Format 'yyyyMMdd').csv"
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            Export-AuctionReport -OutputPath $saveDialog.FileName
            
            [System.Windows.Forms.MessageBox]::Show(
                "Report exported successfully!`n`nFile: $($saveDialog.FileName)",
                "Export Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    })
    
    # Close
    $btnClose.Add_Click({
        $form.Close()
    })
    
    # Double-click to view
    $dgv.Add_CellDoubleClick({
        $btnView.PerformClick()
    })
    
    #endregion
    
    $form.ShowDialog() | Out-Null
}
#endregion

#region Main Execution
Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║           AUCTION QUOTE TRACKER v1.0                  ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

Initialize-TrackerData
New-TrackerGUI
#endregion
