<#
.SYNOPSIS
    Logistics Analytics Dashboard - Freight Cost & Performance Tracking
    
.DESCRIPTION
    Comprehensive analytics dashboard showing freight costs over time, carrier performance
    metrics, trends analysis, and cost optimization insights. Integrates with Quote Tracker
    and Recipient Manager data.
    
.EXAMPLE
    .\Logistics-Analytics-Dashboard.ps1
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 2025-01-08
    Version: 1.0.0
    Change Date: 
    Change Purpose: Initial Release - Analytics & Performance Dashboard
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms.DataVisualization

$script:AuctionsFile = ".\Data\AuctionQuotes.json"
$script:RecipientsFile = ".\Data\FreightRecipients.json"
$script:Auctions = @()
$script:Recipients = @()

#region Data Loading Functions
function Initialize-AnalyticsData {
    <#
    .SYNOPSIS
        Loads auction and recipient data for analysis
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-08
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    if (Test-Path $script:AuctionsFile) {
        $script:Auctions = Get-Content $script:AuctionsFile -Raw | ConvertFrom-Json
        Write-Host "✅ Loaded $($script:Auctions.Count) auction records" -ForegroundColor Green
    }
    else {
        Write-Host "⚠️ No auction data found. Run Quote Tracker first." -ForegroundColor Yellow
        $script:Auctions = @()
    }
    
    if (Test-Path $script:RecipientsFile) {
        $script:Recipients = Get-Content $script:RecipientsFile -Raw | ConvertFrom-Json
        Write-Host "✅ Loaded $($script:Recipients.Count) freight companies" -ForegroundColor Green
    }
    else {
        Write-Host "⚠️ No recipient data found." -ForegroundColor Yellow
        $script:Recipients = @()
    }
}

function Get-FreightCostTrends {
    <#
    .SYNOPSIS
        Calculates freight cost trends over time
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-08
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    $trends = @{
        MonthlyData = @()
        QuarterlyData = @()
        YearlyData = @()
        TotalSpent = 0
        AverageCost = 0
        LowestCost = 99999
        HighestCost = 0
    }
    
    # Group by month
    $monthlyGroups = $script:Auctions | Where-Object { 
        $_.SelectedQuote -and $_.SelectedQuote.Amount -and $_.AuctionDate 
    } | Group-Object { 
        try {
            ([DateTime]::Parse($_.AuctionDate)).ToString("yyyy-MM")
        } catch {
            "Unknown"
        }
    } | Where-Object { $_.Name -ne "Unknown" }
    
    foreach ($group in $monthlyGroups) {
        $monthTotal = ($group.Group.SelectedQuote.Amount | Measure-Object -Sum).Sum
        $monthAvg = ($group.Group.SelectedQuote.Amount | Measure-Object -Average).Average
        $monthCount = $group.Count
        
        $trends.MonthlyData += @{
            Month = $group.Name
            Total = $monthTotal
            Average = $monthAvg
            Count = $monthCount
        }
        
        $trends.TotalSpent += $monthTotal
    }
    
    # Calculate overall statistics
    $allCosts = $script:Auctions | Where-Object { 
        $_.SelectedQuote -and $_.SelectedQuote.Amount 
    } | ForEach-Object { $_.SelectedQuote.Amount }
    
    if ($allCosts.Count -gt 0) {
        $trends.AverageCost = ($allCosts | Measure-Object -Average).Average
        $trends.LowestCost = ($allCosts | Measure-Object -Minimum).Minimum
        $trends.HighestCost = ($allCosts | Measure-Object -Maximum).Maximum
    }
    
    # Sort monthly data by date
    $trends.MonthlyData = $trends.MonthlyData | Sort-Object Month
    
    return $trends
}

function Get-CarrierPerformanceMetrics {
    <#
    .SYNOPSIS
        Calculates performance metrics for each carrier
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-08
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    $carrierMetrics = @{}
    
    # Analyze each auction's quotes
    foreach ($auction in $script:Auctions) {
        if (!$auction.Quotes -or $auction.Quotes.Count -eq 0) { continue }
        
        foreach ($quote in $auction.Quotes) {
            $company = $quote.Company
            if (!$company) { continue }
            
            if (!$carrierMetrics.ContainsKey($company)) {
                $carrierMetrics[$company] = @{
                    CompanyName = $company
                    TotalQuotes = 0
                    TimesSelected = 0
                    WinRate = 0
                    TotalQuotedAmount = 0
                    AverageQuote = 0
                    LowestQuote = 99999
                    HighestQuote = 0
                    AverageResponseTime = 0
                    ResponseTimes = @()
                    LastUsed = $null
                }
            }
            
            $metrics = $carrierMetrics[$company]
            $metrics.TotalQuotes++
            $metrics.TotalQuotedAmount += $quote.Amount
            
            if ($quote.Amount -lt $metrics.LowestQuote) {
                $metrics.LowestQuote = $quote.Amount
            }
            if ($quote.Amount -gt $metrics.HighestQuote) {
                $metrics.HighestQuote = $quote.Amount
            }
            
            # Check if this quote was selected
            if ($auction.SelectedQuote -and $auction.SelectedQuote.Company -eq $company) {
                $metrics.TimesSelected++
                $metrics.LastUsed = $auction.AuctionDate
            }
            
            # Calculate response time if dates available
            if ($auction.CreatedDate -and $quote.ReceivedDate) {
                try {
                    $created = [DateTime]::Parse($auction.CreatedDate)
                    $received = [DateTime]::Parse($quote.ReceivedDate)
                    $responseHours = ($received - $created).TotalHours
                    if ($responseHours -gt 0 -and $responseHours -lt 720) { # Max 30 days
                        $metrics.ResponseTimes += $responseHours
                    }
                } catch { }
            }
        }
    }
    
    # Calculate final metrics
    foreach ($company in $carrierMetrics.Keys) {
        $metrics = $carrierMetrics[$company]
        
        if ($metrics.TotalQuotes -gt 0) {
            $metrics.WinRate = [math]::Round(($metrics.TimesSelected / $metrics.TotalQuotes) * 100, 2)
            $metrics.AverageQuote = [math]::Round($metrics.TotalQuotedAmount / $metrics.TotalQuotes, 2)
        }
        
        if ($metrics.ResponseTimes.Count -gt 0) {
            $metrics.AverageResponseTime = [math]::Round(($metrics.ResponseTimes | Measure-Object -Average).Average, 1)
        }
        
        # Update recipient data if available
        $recipient = $script:Recipients | Where-Object { $_.CompanyName -eq $company }
        if ($recipient) {
            $recipient.TimesUsed = $metrics.TimesSelected
            $recipient.LastUsed = $metrics.LastUsed
        }
    }
    
    return $carrierMetrics.Values | Sort-Object TimesSelected -Descending
}

function Get-CostOptimizationInsights {
    <#
    .SYNOPSIS
        Provides cost optimization insights and recommendations
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-08
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    $insights = @{
        Recommendations = @()
        Warnings = @()
        Opportunities = @()
    }
    
    # Analyze carrier pricing patterns
    $carrierMetrics = Get-CarrierPerformanceMetrics
    
    # Find most cost-effective carriers
    $topPerformers = $carrierMetrics | Where-Object { $_.TimesSelected -gt 0 } | 
        Sort-Object AverageQuote | Select-Object -First 3
    
    if ($topPerformers) {
        $insights.Recommendations += "Top cost-effective carriers: $(($topPerformers.CompanyName) -join ', ')"
    }
    
    # Identify underutilized low-cost carriers
    $underutilized = $carrierMetrics | Where-Object { 
        $_.TotalQuotes -gt 2 -and $_.WinRate -lt 20 -and $_.AverageQuote -lt ($carrierMetrics.AverageQuote | Measure-Object -Average).Average
    }
    
    if ($underutilized) {
        foreach ($carrier in $underutilized) {
            $insights.Opportunities += "Consider using $($carrier.CompanyName) more - avg quote: `$$($carrier.AverageQuote) (win rate: $($carrier.WinRate)%)"
        }
    }
    
    # Check for price increases
    $trends = Get-FreightCostTrends
    if ($trends.MonthlyData.Count -ge 3) {
        $recent = $trends.MonthlyData | Select-Object -Last 3
        $older = $trends.MonthlyData | Select-Object -First 3
        
        $recentAvg = ($recent.Average | Measure-Object -Average).Average
        $olderAvg = ($older.Average | Measure-Object -Average).Average
        
        if ($recentAvg -gt $olderAvg * 1.15) {
            $percentIncrease = [math]::Round((($recentAvg - $olderAvg) / $olderAvg) * 100, 1)
            $insights.Warnings += "Freight costs increased $percentIncrease% compared to earlier period"
        }
    }
    
    # Identify slow responders
    $slowResponders = $carrierMetrics | Where-Object { 
        $_.AverageResponseTime -gt 48 
    } | Sort-Object AverageResponseTime -Descending
    
    if ($slowResponders) {
        $insights.Warnings += "Slow response times from: $(($slowResponders | Select-Object -First 3).CompanyName -join ', ')"
    }
    
    return $insights
}
#endregion

#region Chart Functions
function New-FreightCostChart {
    <#
    .SYNOPSIS
        Creates a line chart showing freight costs over time
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-08
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param(
        [object]$ChartControl,
        [array]$MonthlyData
    )
    
    $ChartControl.Series.Clear()
    $ChartControl.ChartAreas.Clear()
    
    # Create chart area
    $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartArea.Name = "MainArea"
    $chartArea.AxisX.Title = "Month"
    $chartArea.AxisY.Title = "Cost ($)"
    $chartArea.AxisX.Interval = 1
    $chartArea.AxisY.LabelStyle.Format = "C0"
    $chartArea.BackColor = [System.Drawing.Color]::White
    $ChartControl.ChartAreas.Add($chartArea)
    
    # Create series for total costs
    $seriesTotal = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $seriesTotal.Name = "Total Cost"
    $seriesTotal.ChartType = "Line"
    $seriesTotal.BorderWidth = 3
    $seriesTotal.Color = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $seriesTotal.MarkerStyle = "Circle"
    $seriesTotal.MarkerSize = 8
    
    # Create series for average costs
    $seriesAvg = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $seriesAvg.Name = "Average Cost"
    $seriesAvg.ChartType = "Line"
    $seriesAvg.BorderWidth = 2
    $seriesAvg.Color = [System.Drawing.Color]::FromArgb(46, 204, 113)
    $seriesAvg.MarkerStyle = "Square"
    $seriesAvg.MarkerSize = 6
    $seriesAvg.BorderDashStyle = "Dash"
    
    # Add data points
    foreach ($month in $MonthlyData) {
        $monthLabel = ([DateTime]::ParseExact($month.Month, "yyyy-MM", $null)).ToString("MMM yy")
        
        $pointTotal = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
        $pointTotal.SetValueXY($monthLabel, $month.Total)
        $pointTotal.ToolTip = "Total: `$$($month.Total.ToString('N2'))`nCount: $($month.Count)"
        $seriesTotal.Points.Add($pointTotal)
        
        $pointAvg = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
        $pointAvg.SetValueXY($monthLabel, $month.Average)
        $pointAvg.ToolTip = "Average: `$$($month.Average.ToString('N2'))"
        $seriesAvg.Points.Add($pointAvg)
    }
    
    $ChartControl.Series.Add($seriesTotal)
    $ChartControl.Series.Add($seriesAvg)
    
    # Add legend
    $legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $legend.Name = "Legend"
    $legend.Docking = "Bottom"
    $ChartControl.Legends.Add($legend)
}

function New-CarrierPerformanceChart {
    <#
    .SYNOPSIS
        Creates a bar chart showing carrier win rates
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-08
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    param(
        [object]$ChartControl,
        [array]$CarrierMetrics
    )
    
    $ChartControl.Series.Clear()
    $ChartControl.ChartAreas.Clear()
    
    # Create chart area
    $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartArea.Name = "MainArea"
    $chartArea.AxisX.Title = "Carrier"
    $chartArea.AxisY.Title = "Win Rate (%)"
    $chartArea.AxisX.Interval = 1
    $chartArea.BackColor = [System.Drawing.Color]::White
    $ChartControl.ChartAreas.Add($chartArea)
    
    # Create series
    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series
    $series.Name = "Win Rate"
    $series.ChartType = "Column"
    $series.Color = [System.Drawing.Color]::FromArgb(155, 89, 182)
    
    # Add top 10 carriers by selection count
    $topCarriers = $CarrierMetrics | Where-Object { $_.TotalQuotes -gt 0 } | 
        Sort-Object TimesSelected -Descending | Select-Object -First 10
    
    foreach ($carrier in $topCarriers) {
        $point = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
        $point.SetValueXY($carrier.CompanyName, $carrier.WinRate)
        $point.ToolTip = "Win Rate: $($carrier.WinRate)%`nSelected: $($carrier.TimesSelected)/$($carrier.TotalQuotes)"
        $point.Color = if ($carrier.WinRate -ge 50) { 
            [System.Drawing.Color]::FromArgb(46, 204, 113) 
        } elseif ($carrier.WinRate -ge 25) { 
            [System.Drawing.Color]::FromArgb(241, 196, 15) 
        } else { 
            [System.Drawing.Color]::FromArgb(231, 76, 60) 
        }
        $series.Points.Add($point)
    }
    
    $ChartControl.Series.Add($series)
}
#endregion

#region Dashboard GUI
function New-AnalyticsDashboard {
    <#
    .SYNOPSIS
        Creates the analytics dashboard GUI
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 2025-01-08
        Version: 1.0.0
        Change Date: 
        Change Purpose:
    #>
    
    # Main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Logistics Analytics Dashboard"
    $form.Size = New-Object System.Drawing.Size(1400, 900)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = [System.Drawing.Color]::WhiteSmoke
    
    # Header Panel
    $pnlHeader = New-Object System.Windows.Forms.Panel
    $pnlHeader.Location = New-Object System.Drawing.Point(0, 0)
    $pnlHeader.Size = New-Object System.Drawing.Size(1400, 80)
    $pnlHeader.BackColor = [System.Drawing.Color]::FromArgb(44, 62, 80)
    
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    $lblTitle.Size = New-Object System.Drawing.Size(600, 30)
    $lblTitle.Text = "📊 Logistics Analytics Dashboard"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $pnlHeader.Controls.Add($lblTitle)
    
    $lblSubtitle = New-Object System.Windows.Forms.Label
    $lblSubtitle.Location = New-Object System.Drawing.Point(20, 50)
    $lblSubtitle.Size = New-Object System.Drawing.Size(600, 20)
    $lblSubtitle.Text = "Freight cost trends & carrier performance metrics"
    $lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblSubtitle.ForeColor = [System.Drawing.Color]::LightGray
    $pnlHeader.Controls.Add($lblSubtitle)
    
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Location = New-Object System.Drawing.Point(1250, 25)
    $btnRefresh.Size = New-Object System.Drawing.Size(120, 35)
    $btnRefresh.Text = "🔄 Refresh"
    $btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    $btnRefresh.ForeColor = [System.Drawing.Color]::White
    $btnRefresh.FlatStyle = "Flat"
    $btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $pnlHeader.Controls.Add($btnRefresh)
    
    $form.Controls.Add($pnlHeader)
    
    # Tab Control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 90)
    $tabControl.Size = New-Object System.Drawing.Size(1365, 750)
    
    #region Overview Tab
    $tabOverview = New-Object System.Windows.Forms.TabPage
    $tabOverview.Text = "Overview"
    
    # Statistics Cards
    $yPos = 20
    
    # Total Spent Card
    $pnlTotalSpent = New-Object System.Windows.Forms.Panel
    $pnlTotalSpent.Location = New-Object System.Drawing.Point(20, $yPos)
    $pnlTotalSpent.Size = New-Object System.Drawing.Size(300, 100)
    $pnlTotalSpent.BackColor = [System.Drawing.Color]::FromArgb(52, 152, 219)
    
    $lblTotalSpentTitle = New-Object System.Windows.Forms.Label
    $lblTotalSpentTitle.Location = New-Object System.Drawing.Point(15, 15)
    $lblTotalSpentTitle.Size = New-Object System.Drawing.Size(270, 20)
    $lblTotalSpentTitle.Text = "Total Freight Costs"
    $lblTotalSpentTitle.ForeColor = [System.Drawing.Color]::White
    $lblTotalSpentTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $pnlTotalSpent.Controls.Add($lblTotalSpentTitle)
    
    $lblTotalSpent = New-Object System.Windows.Forms.Label
    $lblTotalSpent.Location = New-Object System.Drawing.Point(15, 40)
    $lblTotalSpent.Size = New-Object System.Drawing.Size(270, 40)
    $lblTotalSpent.Text = "$0.00"
    $lblTotalSpent.ForeColor = [System.Drawing.Color]::White
    $lblTotalSpent.Font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
    $pnlTotalSpent.Controls.Add($lblTotalSpent)
    
    $tabOverview.Controls.Add($pnlTotalSpent)
    
    # Average Cost Card
    $pnlAvgCost = New-Object System.Windows.Forms.Panel
    $pnlAvgCost.Location = New-Object System.Drawing.Point(340, $yPos)
    $pnlAvgCost.Size = New-Object System.Drawing.Size(300, 100)
    $pnlAvgCost.BackColor = [System.Drawing.Color]::FromArgb(46, 204, 113)
    
    $lblAvgCostTitle = New-Object System.Windows.Forms.Label
    $lblAvgCostTitle.Location = New-Object System.Drawing.Point(15, 15)
    $lblAvgCostTitle.Size = New-Object System.Drawing.Size(270, 20)
    $lblAvgCostTitle.Text = "Average Cost per Shipment"
    $lblAvgCostTitle.ForeColor = [System.Drawing.Color]::White
    $lblAvgCostTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $pnlAvgCost.Controls.Add($lblAvgCostTitle)
    
    $lblAvgCost = New-Object System.Windows.Forms.Label
    $lblAvgCost.Location = New-Object System.Drawing.Point(15, 40)
    $lblAvgCost.Size = New-Object System.Drawing.Size(270, 40)
    $lblAvgCost.Text = "$0.00"
    $lblAvgCost.ForeColor = [System.Drawing.Color]::White
    $lblAvgCost.Font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
    $pnlAvgCost.Controls.Add($lblAvgCost)
    
    $tabOverview.Controls.Add($pnlAvgCost)
    
    # Shipment Count Card
    $pnlShipments = New-Object System.Windows.Forms.Panel
    $pnlShipments.Location = New-Object System.Drawing.Point(660, $yPos)
    $pnlShipments.Size = New-Object System.Drawing.Size(300, 100)
    $pnlShipments.BackColor = [System.Drawing.Color]::FromArgb(155, 89, 182)
    
    $lblShipmentsTitle = New-Object System.Windows.Forms.Label
    $lblShipmentsTitle.Location = New-Object System.Drawing.Point(15, 15)
    $lblShipmentsTitle.Size = New-Object System.Drawing.Size(270, 20)
    $lblShipmentsTitle.Text = "Total Shipments"
    $lblShipmentsTitle.ForeColor = [System.Drawing.Color]::White
    $lblShipmentsTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $pnlShipments.Controls.Add($lblShipmentsTitle)
    
    $lblShipments = New-Object System.Windows.Forms.Label
    $lblShipments.Location = New-Object System.Drawing.Point(15, 40)
    $lblShipments.Size = New-Object System.Drawing.Size(270, 40)
    $lblShipments.Text = "0"
    $lblShipments.ForeColor = [System.Drawing.Color]::White
    $lblShipments.Font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
    $pnlShipments.Controls.Add($lblShipments)
    
    $tabOverview.Controls.Add($pnlShipments)
    
    # Carriers Card
    $pnlCarriers = New-Object System.Windows.Forms.Panel
    $pnlCarriers.Location = New-Object System.Drawing.Point(980, $yPos)
    $pnlCarriers.Size = New-Object System.Drawing.Size(300, 100)
    $pnlCarriers.BackColor = [System.Drawing.Color]::FromArgb(230, 126, 34)
    
    $lblCarriersTitle = New-Object System.Windows.Forms.Label
    $lblCarriersTitle.Location = New-Object System.Drawing.Point(15, 15)
    $lblCarriersTitle.Size = New-Object System.Drawing.Size(270, 20)
    $lblCarriersTitle.Text = "Active Carriers"
    $lblCarriersTitle.ForeColor = [System.Drawing.Color]::White
    $lblCarriersTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $pnlCarriers.Controls.Add($lblCarriersTitle)
    
    $lblCarriers = New-Object System.Windows.Forms.Label
    $lblCarriers.Location = New-Object System.Drawing.Point(15, 40)
    $lblCarriers.Size = New-Object System.Drawing.Size(270, 40)
    $lblCarriers.Text = "0"
    $lblCarriers.ForeColor = [System.Drawing.Color]::White
    $lblCarriers.Font = New-Object System.Drawing.Font("Segoe UI", 24, [System.Drawing.FontStyle]::Bold)
    $pnlCarriers.Controls.Add($lblCarriers)
    
    $tabOverview.Controls.Add($pnlCarriers)
    
    $yPos += 120
    
    # Freight Cost Trend Chart
    $lblChartTitle = New-Object System.Windows.Forms.Label
    $lblChartTitle.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblChartTitle.Size = New-Object System.Drawing.Size(400, 25)
    $lblChartTitle.Text = "Freight Cost Trends"
    $lblChartTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $tabOverview.Controls.Add($lblChartTitle)
    
    $yPos += 35
    
    $chartCostTrend = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chartCostTrend.Location = New-Object System.Drawing.Point(20, $yPos)
    $chartCostTrend.Size = New-Object System.Drawing.Size(1300, 350)
    $chartCostTrend.BackColor = [System.Drawing.Color]::White
    $tabOverview.Controls.Add($chartCostTrend)
    
    $yPos += 370
    
    # Insights Section
    $lblInsights = New-Object System.Windows.Forms.Label
    $lblInsights.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblInsights.Size = New-Object System.Drawing.Size(400, 25)
    $lblInsights.Text = "💡 Cost Optimization Insights"
    $lblInsights.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $tabOverview.Controls.Add($lblInsights)
    
    $yPos += 35
    
    $txtInsights = New-Object System.Windows.Forms.TextBox
    $txtInsights.Location = New-Object System.Drawing.Point(20, $yPos)
    $txtInsights.Size = New-Object System.Drawing.Size(1300, 100)
    $txtInsights.Multiline = $true
    $txtInsights.ScrollBars = "Vertical"
    $txtInsights.ReadOnly = $true
    $txtInsights.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $tabOverview.Controls.Add($txtInsights)
    
    $tabControl.TabPages.Add($tabOverview)
    #endregion
    
    #region Carrier Performance Tab
    $tabCarriers = New-Object System.Windows.Forms.TabPage
    $tabCarriers.Text = "Carrier Performance"
    
    $yPos = 20
    
    # Performance Chart
    $lblPerfChart = New-Object System.Windows.Forms.Label
    $lblPerfChart.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblPerfChart.Size = New-Object System.Drawing.Size(400, 25)
    $lblPerfChart.Text = "Carrier Win Rates"
    $lblPerfChart.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $tabCarriers.Controls.Add($lblPerfChart)
    
    $yPos += 35
    
    $chartCarrierPerf = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
    $chartCarrierPerf.Location = New-Object System.Drawing.Point(20, $yPos)
    $chartCarrierPerf.Size = New-Object System.Drawing.Size(1300, 300)
    $chartCarrierPerf.BackColor = [System.Drawing.Color]::White
    $tabCarriers.Controls.Add($chartCarrierPerf)
    
    $yPos += 320
    
    # Carrier Details Grid
    $lblCarrierGrid = New-Object System.Windows.Forms.Label
    $lblCarrierGrid.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblCarrierGrid.Size = New-Object System.Drawing.Size(400, 25)
    $lblCarrierGrid.Text = "Detailed Carrier Metrics"
    $lblCarrierGrid.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $tabCarriers.Controls.Add($lblCarrierGrid)
    
    $yPos += 35
    
    $dgvCarriers = New-Object System.Windows.Forms.DataGridView
    $dgvCarriers.Location = New-Object System.Drawing.Point(20, $yPos)
    $dgvCarriers.Size = New-Object System.Drawing.Size(1300, 300)
    $dgvCarriers.AllowUserToAddRows = $false
    $dgvCarriers.AllowUserToDeleteRows = $false
    $dgvCarriers.ReadOnly = $true
    $dgvCarriers.SelectionMode = "FullRowSelect"
    $dgvCarriers.AutoSizeColumnsMode = "Fill"
    
    $dgvCarriers.Columns.Add("CompanyName", "Carrier") | Out-Null
    $dgvCarriers.Columns.Add("TotalQuotes", "Quotes") | Out-Null
    $dgvCarriers.Columns.Add("TimesSelected", "Selected") | Out-Null
    $dgvCarriers.Columns.Add("WinRate", "Win Rate %") | Out-Null
    $dgvCarriers.Columns.Add("AverageQuote", "Avg Quote") | Out-Null
    $dgvCarriers.Columns.Add("ResponseTime", "Avg Response (hrs)") | Out-Null
    
    $dgvCarriers.Columns["TotalQuotes"].Width = 80
    $dgvCarriers.Columns["TimesSelected"].Width = 80
    $dgvCarriers.Columns["WinRate"].Width = 100
    $dgvCarriers.Columns["AverageQuote"].Width = 100
    $dgvCarriers.Columns["ResponseTime"].Width = 150
    
    $tabCarriers.Controls.Add($dgvCarriers)
    
    $tabControl.TabPages.Add($tabCarriers)
    #endregion
    
    $form.Controls.Add($tabControl)
    
    #region Helper Functions
    function Update-Dashboard {
        # Get analytics data
        $trends = Get-FreightCostTrends
        $carrierMetrics = Get-CarrierPerformanceMetrics
        $insights = Get-CostOptimizationInsights
        
        # Update overview cards
        $lblTotalSpent.Text = "$($trends.TotalSpent.ToString('C2'))"
        $lblAvgCost.Text = "$($trends.AverageCost.ToString('C2'))"
        
        $shipmentsWithCost = ($script:Auctions | Where-Object { 
            $_.SelectedQuote -and $_.SelectedQuote.Amount 
        }).Count
        $lblShipments.Text = $shipmentsWithCost.ToString()
        
        $activeCarriers = ($carrierMetrics | Where-Object { $_.TimesSelected -gt 0 }).Count
        $lblCarriers.Text = $activeCarriers.ToString()
        
        # Update cost trend chart
        if ($trends.MonthlyData.Count -gt 0) {
            New-FreightCostChart -ChartControl $chartCostTrend -MonthlyData $trends.MonthlyData
        }
        
        # Update carrier performance chart
        if ($carrierMetrics.Count -gt 0) {
            New-CarrierPerformanceChart -ChartControl $chartCarrierPerf -CarrierMetrics $carrierMetrics
        }
        
        # Update carrier grid
        $dgvCarriers.Rows.Clear()
        foreach ($carrier in $carrierMetrics) {
            $respTimeDisplay = if ($carrier.AverageResponseTime -gt 0) {
                "$($carrier.AverageResponseTime) hrs"
            } else {
                'N/A'
            }

            $row = @(
                $carrier.CompanyName,
                $carrier.TotalQuotes,
                $carrier.TimesSelected,
                "$($carrier.WinRate)%",
                "$($carrier.AverageQuote.ToString('C2'))",
                $respTimeDisplay
            )
        }
        
        # Update insights
        $insightText = ""
        
        if ($insights.Recommendations.Count -gt 0) {
            $insightText += "✅ RECOMMENDATIONS:`n"
            foreach ($rec in $insights.Recommendations) {
                $insightText += "  • $rec`n"
            }
            $insightText += "`n"
        }
        
        if ($insights.Opportunities.Count -gt 0) {
            $insightText += "💰 OPPORTUNITIES:`n"
            foreach ($opp in $insights.Opportunities) {
                $insightText += "  • $opp`n"
            }
            $insightText += "`n"
        }
        
        if ($insights.Warnings.Count -gt 0) {
            $insightText += "⚠️ WARNINGS:`n"
            foreach ($warn in $insights.Warnings) {
                $insightText += "  • $warn`n"
            }
        }
        
        if (!$insightText) {
            $insightText = "No insights available yet. Track more auctions to see recommendations."
        }
        
        $txtInsights.Text = $insightText
    }
    #endregion
    
    #region Event Handlers
    $btnRefresh.Add_Click({
        Initialize-AnalyticsData
        Update-Dashboard
        
        [System.Windows.Forms.MessageBox]::Show(
            "Dashboard refreshed successfully!",
            "Refresh Complete",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    
    # Initial load
    $form.Add_Shown({
        Update-Dashboard
    })
    #endregion
    
    $form.ShowDialog() | Out-Null
}
#endregion

#region Main Execution
Write-Host "`n╔════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║      LOGISTICS ANALYTICS DASHBOARD v1.0               ║" -ForegroundColor Cyan
Write-Host "╚════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

Initialize-AnalyticsData
New-AnalyticsDashboard
#endregion