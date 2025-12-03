<#
.SYNOPSIS
    Automated Logistics Quote Email Generator

.DESCRIPTION
    Automates the creation of freight-quote request emails, including:
      • Importing auction lot data from CSV
      • Collecting lot images and generating a gallery-style HTML report
      • Converting the HTML report to PDF using a robust, multi-engine pipeline
      • Automatically creating an Outlook draft email with the PDF attached
      • Generating logs, a processing report, and an optional dashboard

    PDF CONVERSION PIPELINE
    ------------------------
    Beginning with version 1.8.3, PDF generation uses a hardened,
    timeout-protected engine with browser and CLI failover:

      1. Microsoft Edge (headless)
      2. Google Chrome (headless)
      3. wkhtmltopdf (if installed)
      4. LibreOffice (if installed)

    Each engine is executed with:
      • A configurable timeout (default 60s)
      • A safe temp workspace to avoid path issues (C:\Temp\JTLS)
      • Hardened chromium flags:
          --print-to-pdf-no-header
          --kiosk-printing
          --allow-file-access-from-files
          --virtual-time-budget=10000

    NEW BEHAVIOR (v1.8.3)
    ----------------------
    If a PDF file is successfully written *even if the engine times out*,
    the script now treats the conversion as SUCCESS. This resolves cases
    where Chrome/Edge writes the PDF but fails to exit in time.

    The PDF is attached to the Outlook draft unless:
      • No converter succeeds AND
      • No PDF file exists
    In that case, the HTML report is attached instead.

.PARAMETER CSVPath
    Path to CSV containing auction lot data.

.PARAMETER ImageDirectory
    Directory containing lot images in jpg format.

.PARAMETER OutputDirectory
    Output path for HTML, PDF, and reports.

.PARAMETER LogDirectory
    Path where logs are written.

.PARAMETER TemplatePath
    Optional HTML email template file.

.PARAMETER ConfigPath
    Optional JSON configuration containing auction metadata.

.PARAMETER MaxImagesPerLot
    Maximum number of images included per lot (default is 3).

.PARAMETER CreateOutlookDraft
    Creates an Outlook draft email with the generated email HTML and PDF attachment.

.PARAMETER ShowDashboard
    Displays an interactive dashboard summarizing processing.

.PARAMETER RequesterName
    Name used in the email signature.

.PARAMETER RequesterPhone
    Phone number used in the email signature.

.EXAMPLE
    .\Generate-LogisticsEmail.ps1 `
        -CSVPath ".\auction.csv" `
        -ImageDirectory ".\Images" `
        -OutputDirectory ".\Output" `
        -CreateOutlookDraft

.EXAMPLE
    .\Generate-LogisticsEmail.ps1 `
        -CSVPath ".\auction.csv" `
        -ImageDirectory ".\Images" `
        -ConfigPath ".\auction.json" `
        -TemplatePath ".\LogisticsEmail_Template.html"

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Version: 1.8.3
    Change Date: 2025-12-02

    CHANGELOG
    ---------
    1.8.3 - Added timeout-aware PDF trust mode, eliminating hangs caused by
            long-running Chromium processes. Strengthened browser flags.
            Hardened CLI timeout handling.

    1.8.2 - Major cleanup. Fixed brace mismatch. No nested functions.
            PSScriptAnalyzer-friendly. Better PDF layout CSS.

    1.8.0 - Replacement of Word PDF engine. Multi-engine conversion pipeline.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$TemplatePath = $null,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CSVPath,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$ImageDirectory,

    [Parameter(Mandatory = $false)]
    [string]$OutputDirectory = ".\Output",

    [Parameter(Mandatory = $false)]
    [string]$LogDirectory = ".\Logs",

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 20)]
    [int]$MaxImagesPerLot = 3,

    [Parameter(Mandatory = $false)]
    [switch]$CreateOutlookDraft,

    [Parameter(Mandatory = $false)]
    [switch]$ShowDashboard,

    [Parameter(Mandatory = $false)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ConfigPath,

    [Parameter(Mandatory = $false)]
    [string]$RequesterName = $env:USERNAME,

    [Parameter(Mandatory = $false)]
    [string]$RequesterPhone = ""
)

#region Logging
function Write-JTLSLog {
    <#
    .SYNOPSIS
        Writes a log entry to console and files.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','DEBUG','SUCCESS')]
        [string]$Level = 'INFO'
    )
    $timeStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$timeStamp] [$Level] $Message"

    switch ($Level) {
        'ERROR'   { Write-Host $entry -ForegroundColor Red }
        'WARNING' { Write-Host $entry -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green }
        'DEBUG'   { Write-Host $entry -ForegroundColor Gray }
        default   { Write-Host $entry -ForegroundColor White }
    }

    if ($script:LogFile)      { Add-Content -Path $script:LogFile -Value $entry }
    if ($Level -eq 'ERROR' -and $script:ErrorLogFile) { Add-Content -Path $script:ErrorLogFile -Value $entry }
}

function Start-JTLSLogging {
    <#
    .SYNOPSIS
        Initializes logging files and runtime statistics.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LogDir)

    if (-not (Test-Path $LogDir)) {
        $null = New-Item -ItemType Directory -Path $LogDir -Force
    }

    $script:LogFile       = Join-Path $LogDir ("LogisticsEmail_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $script:ErrorLogFile  = Join-Path $LogDir ("LogisticsEmail_Errors_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $script:ProcessingStats = @{
        StartTime       = Get-Date
        TotalLots       = 0
        ProcessedLots   = 0
        FailedLots      = 0
        ImagesFound     = 0
        ImagesMissing   = 0
        PDFsGenerated   = 0
        EmailsGenerated = 0
        DataSources     = @()
    }

    Write-JTLSLog -Message '=== Logistics Email Automation Started ===' -Level INFO
    Write-JTLSLog -Message 'Script Version: 1.8.2' -Level INFO
    Write-JTLSLog -Message ("User: {0}" -f $env:USERNAME) -Level INFO
    Write-JTLSLog -Message ("Machine: {0}" -f $env:COMPUTERNAME) -Level INFO
}
#endregion

#region Configuration and Data
function Get-AuctionConfig {
    <#
    .SYNOPSIS
        Loads JSON configuration file if provided.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ConfigFile)

    try {
        Write-JTLSLog -Message ("Loading configuration from: {0}" -f $ConfigFile) -Level INFO
        return (Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json)
    }
    catch {
        Write-JTLSLog -Message "Failed to load configuration: $_" -Level ERROR
        return $null
    }
}

function Import-AuctionData {
    <#
    .SYNOPSIS
        Imports and validates auction data from CSV.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    try {
        Write-JTLSLog -Message ("Importing CSV data from: {0}" -f $Path) -Level INFO
        $script:ProcessingStats.DataSources += $Path
        $data = Import-Csv -Path $Path
        $script:ProcessingStats.TotalLots = $data.Count
        Write-JTLSLog -Message ("Successfully imported {0} lots" -f $data.Count) -Level SUCCESS

        $required = @('Lot','Description')
        $missing = $required | Where-Object { $_ -notin $data[0].PSObject.Properties.Name }
        if ($missing) {
            throw ("Missing required columns: {0}" -f ($missing -join ', '))
        }
        return $data
    }
    catch {
        Write-JTLSLog -Message "Failed to import CSV: $_" -Level ERROR
        throw
    }
}

function Get-LotImage {
    <#
    .SYNOPSIS
        Retrieves the best images for each lot with simple prioritization.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Lots,
        [Parameter(Mandatory)][string]$ImageDir,
        [Parameter(Mandatory)][int]$MaxImagesPerLot
    )

    $results = @()

    foreach ($lot in $Lots) {
        $lotNumber = $lot.Lot
        $all = @()

        # Primary
        $primary = Join-Path $ImageDir ("{0}.jpg" -f $lotNumber)
        if (Test-Path $primary) {
            $all += [pscustomobject]@{ Path = $primary; Priority = 1; SortOrder = 0; Type = 'Primary' }
        }

        # Numbered 2..30
        for ($i = 2; $i -le 30; $i++) {
            $p = Join-Path $ImageDir ("{0}-{1}.jpg" -f $lotNumber, $i)
            if (Test-Path $p) {
                $all += [pscustomobject]@{ Path = $p; Priority = 2; SortOrder = $i; Type = ("Image{0}" -f $i) }
            }
        }

        # Variants a..f
        foreach ($letter in 'a','b','c','d','e','f') {
            $p = Join-Path $ImageDir ("{0}-{1}.jpg" -f $lotNumber, $letter)
            if (Test-Path $p) {
                $all += [pscustomobject]@{ Path = $p; Priority = 3; SortOrder = (100 + [int][char]$letter); Type = ("Variant-{0}" -f $letter) }
            }
        }

        if ($all.Count -gt 0) {
            $selected = $all | Sort-Object Priority, SortOrder | Select-Object -First $MaxImagesPerLot
            $imgPaths = $selected.Path
            $script:ProcessingStats.ImagesFound += $imgPaths.Count

            $results += [pscustomobject]@{
                LotNumber   = $lotNumber
                Description = $lot.Description
                ImagePaths  = $imgPaths
                ImageCount  = $imgPaths.Count
                TotalFound  = $all.Count
                FileSize    = ($imgPaths | ForEach-Object { (Get-Item $_).Length } | Measure-Object -Sum).Sum
            }
        }
        else {
            Write-JTLSLog -Message ("No images found for Lot {0}" -f $lotNumber) -Level WARNING
            $script:ProcessingStats.ImagesMissing++
        }
    }

    $included = ($results.ImageCount | Measure-Object -Sum).Sum
    $found = ($results.TotalFound | Measure-Object -Sum).Sum
    Write-JTLSLog -Message ("Including {0} of {1} total images found" -f $included, $found) -Level INFO

    return $results
}
#endregion

#region HTML Image Report
function New-LotHtmlReport {
    <#
    .SYNOPSIS
        Builds a self-contained HTML gallery report from lot images.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][array]$Images,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$AuctionName
    )

    try {
        $totalImageCount = 0
        foreach ($lot in $Images) {
            if ($lot.ImagePaths) {
                $totalImageCount += $lot.ImagePaths.Count
            }
        }

        Write-JTLSLog -Message ("Generating image report for {0} lots with {1} total images" -f $Images.Count, $totalImageCount) -Level INFO
        $reportPath = Join-Path (Resolve-Path $OutputPath).Path ("AuctionLots_{0}.html" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.IO.StreamWriter($reportPath, $false, $utf8NoBom)

        try {
            $writer.WriteLine('<!DOCTYPE html>')
            $writer.WriteLine('<html>')
            $writer.WriteLine('<head>')
            $writer.WriteLine('  <meta charset="UTF-8">')
            $writer.WriteLine('  <title>Auction Lot Images - ' + $AuctionName + '</title>')
            $writer.WriteLine('  <style>')
            $writer.WriteLine('    @page { size: Letter; margin: 0.5in; }')
            $writer.WriteLine('    * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }')
            $writer.WriteLine('    html, body { height: 100%; }')
            $writer.WriteLine('    body { font-family: Arial, sans-serif; margin: 0; background: #f5f5f5; }')
            $writer.WriteLine('    .header { background: #2c3e50; color: white; padding: 20px; text-align: center; margin: 20px; border-radius: 8px; }')
            $writer.WriteLine('    .page { background: white; margin: 20px; padding: 16px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); border-radius: 8px; break-after: page; page-break-after: always; }')
            $writer.WriteLine('    .page:last-child { break-after: auto; page-break-after: auto; }')
            $writer.WriteLine('    .lot-number { font-size: 20px; font-weight: bold; color: #2c3e50; margin-bottom: 6px; border-bottom: 2px solid #3498db; padding-bottom: 4px; }')
            $writer.WriteLine('    .description { font-size: 12px; color: #444; margin-bottom: 10px; font-style: italic; }')
            $writer.WriteLine('    .image-container { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }')
            $writer.WriteLine('    .image-wrapper { background: #f8f9fa; padding: 8px; border-radius: 6px; page-break-inside: avoid; break-inside: avoid; }')
            $writer.WriteLine('    img { display: block; width: 100%; height: auto; border: 1px solid #ddd; border-radius: 4px; box-shadow: 0 1px 2px rgba(0,0,0,0.08); max-height: 4.1in; object-fit: contain; }')
            $writer.WriteLine('    .image-label { text-align: center; font-size: 10px; color: #666; margin-top: 6px; }')
            $writer.WriteLine('    @media print { body { background: white; } .header { margin: 0 0 12px 0; border-radius: 0; box-shadow: none; } .page { margin: 0; box-shadow: none; border-radius: 0; } }')
            $writer.WriteLine('  </style>')
            $writer.WriteLine('</head>')
            $writer.WriteLine('<body>')
            $writer.WriteLine('  <div class="header">')
            $writer.WriteLine('    <h1>Auction Lot Images</h1>')
            $writer.WriteLine('    <div class="stats">')
            $writer.WriteLine('      Generated: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm') + '<br>')
            $writer.WriteLine('      Total Lots: ' + $Images.Count + '<br>')
            $writer.WriteLine('      Total Images: ' + $totalImageCount)
            $writer.WriteLine('    </div>')
            $writer.WriteLine('  </div>')

            foreach ($lotInfo in $Images) {
                $writer.WriteLine('  <div class="page">')
                $writer.WriteLine('    <div class="lot-number">Lot #' + $lotInfo.LotNumber + '</div>')
                $writer.WriteLine('    <div class="description">' + $lotInfo.Description + '</div>')
                $writer.WriteLine('    <div class="image-container">')

                if ($lotInfo.ImagePaths -and $lotInfo.ImagePaths.Count -gt 0) {
                    $imageNum = 1
                    $totalForLot = $lotInfo.ImagePaths.Count

                    foreach ($imagePath in $lotInfo.ImagePaths) {
                        if ($imagePath -and (Test-Path $imagePath)) {
                            try {
                                $imageBytes  = [System.IO.File]::ReadAllBytes($imagePath)
                                $imageBase64 = [System.Convert]::ToBase64String($imageBytes)
                                $imageSrc    = 'data:image/jpeg;base64,' + $imageBase64
                                $label       = if ($totalForLot -gt 1) { "Image $imageNum of $totalForLot" } else { "Lot Image" }

                                $writer.WriteLine('      <div class="image-wrapper">')
                                $writer.WriteLine('        <img src="' + $imageSrc + '" alt="Lot ' + $lotInfo.LotNumber + ' - ' + $label + '" />')
                                $writer.WriteLine('        <div class="image-label">' + $label + '</div>')
                                $writer.WriteLine('      </div>')

                                $imageNum++
                            }
                            catch {
                                Write-JTLSLog -Message ("Could not embed image: {0} - {1}" -f $imagePath, $_) -Level WARNING
                            }
                        }
                    }
                }
                else {
                    $writer.WriteLine('      <div class="image-wrapper">')
                    $writer.WriteLine('        <div style="padding: 20px; background: #f0f0f0; text-align: center; border-radius: 5px;">')
                    $writer.WriteLine('          <p style="color: #666; margin: 0;">No images available for this lot</p>')
                    $writer.WriteLine('        </div>')
                    $writer.WriteLine('      </div>')
                }

                $writer.WriteLine('    </div>')   # .image-container
                $writer.WriteLine('  </div>')     # .page
            }

            $writer.WriteLine('</body>')
            $writer.WriteLine('</html>')
        }
        finally {
            if ($writer) { $writer.Close(); $writer.Dispose() }
        }

        Write-JTLSLog -Message ("Image report generated: {0}" -f $reportPath) -Level SUCCESS
        $script:ProcessingStats.PDFsGenerated++
        return $reportPath
    }
    catch {
        Write-JTLSLog -Message "Failed to generate image report: $_" -Level ERROR
        throw
    }
}
#endregion

#region HTML → PDF helpers
function Invoke-JTLSCli {
    <#
    .SYNOPSIS
        Invokes a CLI tool and returns $true if the expected output PDF exists.
        Includes a timeout and trusts the PDF if it was created even on timeout.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Exe,
        [Parameter(Mandatory)][string]$Arguments,
        [Parameter(Mandatory)][string]$ExpectedPdf,
        [Parameter(Mandatory = $false)]
        [ValidateRange(5, 300)]
        [int]$TimeoutSeconds = 60
    )

    if (-not (Test-Path $Exe)) {
        return $false
    }

    $outDir = [System.IO.Path]::GetDirectoryName($ExpectedPdf)
    if (-not (Test-Path $outDir)) {
        $null = New-Item -ItemType Directory -Force -Path $outDir
    }

    Write-Host "→ $Exe $Arguments" -ForegroundColor DarkGray

    $process = Start-Process -FilePath $Exe -ArgumentList $Arguments -NoNewWindow -PassThru

    try {
        $waitResult = Wait-Process -Id $process.Id -Timeout $TimeoutSeconds -ErrorAction SilentlyContinue

        $pdfExists = Test-Path $ExpectedPdf

        if (-not $waitResult) {
            if ($pdfExists) {
                Write-Host ("Timed out after {0}s waiting for {1}, but PDF exists at {2}. Treating as success." -f $TimeoutSeconds, (Split-Path $Exe -Leaf), $ExpectedPdf) -ForegroundColor Yellow
                try {
                    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host ("Failed to stop process {0}: {1}" -f $process.Id, $_) -ForegroundColor Red
                }
                return $true
            }

            Write-Host ("Timed out after {0}s waiting for {1}. Killing process..." -f $TimeoutSeconds, (Split-Path $Exe -Leaf)) -ForegroundColor Yellow
            try {
                Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Host ("Failed to stop process {0}: {1}" -f $process.Id, $_) -ForegroundColor Red
            }
            return $false
        }

        if ($process.ExitCode -ne 0) {
            Write-Host ("ExitCode: {0}" -f $process.ExitCode) -ForegroundColor Yellow
        }

        return $pdfExists
    }
    catch {
        Write-Host ("Error while waiting for {0}: {1}" -f (Split-Path $Exe -Leaf), $_) -ForegroundColor Red
        return $false
    }
}

function Invoke-JTLSBrowser {
    <#
    .SYNOPSIS
        Runs a Chromium-based browser with headless print flags and checks for the output PDF.
        Includes a timeout to prevent the script from hanging. If the PDF exists even after a
        timeout, we treat it as success.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Exe,
        [Parameter(Mandatory)][string]$Flags,
        [Parameter(Mandatory)][string]$ExpectedPdf,
        [Parameter(Mandatory = $false)]
        [ValidateRange(5, 300)]
        [int]$TimeoutSeconds = 60
    )

    if (-not (Test-Path $Exe)) {
        return $false
    }

    Write-Host "→ $Exe $Flags" -ForegroundColor DarkGray

    # Start without -Wait so we can control timeout
    $process = Start-Process -FilePath $Exe -ArgumentList $Flags -NoNewWindow -PassThru

    try {
        $waitResult = Wait-Process -Id $process.Id -Timeout $TimeoutSeconds -ErrorAction SilentlyContinue

        # Always check if the PDF exists, even if we timed out
        $pdfExists = Test-Path $ExpectedPdf

        if (-not $waitResult) {
            if ($pdfExists) {
                Write-Host ("Timed out after {0}s waiting for {1}, but PDF exists at {2}. Treating as success." -f $TimeoutSeconds, (Split-Path $Exe -Leaf), $ExpectedPdf) -ForegroundColor Yellow
                try {
                    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-Host ("Failed to stop process {0}: {1}" -f $process.Id, $_) -ForegroundColor Red
                }
                return $true
            }

            Write-Host ("Timed out after {0}s waiting for {1}. Killing process..." -f $TimeoutSeconds, (Split-Path $Exe -Leaf)) -ForegroundColor Yellow
            try {
                Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Host ("Failed to stop process {0}: {1}" -f $process.Id, $_) -ForegroundColor Red
            }
            return $false
        }

        if ($process.ExitCode -ne 0) {
            Write-Host ("ExitCode: {0}" -f $process.ExitCode) -ForegroundColor Yellow
        }

        return $pdfExists
    }
    catch {
        Write-Host ("Error while waiting for {0}: {1}" -f (Split-Path $Exe -Leaf), $_) -ForegroundColor Red
        return $false
    }
}

function Convert-ToPdf {
    <#
    .SYNOPSIS
        Converts HTML to PDF using Edge/Chrome/wkhtmltopdf/LibreOffice with safe temp paths and print-friendly flags.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$HtmlPath,
        [Parameter(Mandatory)][string]$PdfPath
    )

    $pdfFull = [IO.Path]::GetFullPath($PdfPath)
    $null = New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($pdfFull)) -ErrorAction SilentlyContinue

    # Safe working folder to avoid apostrophes/spaces in user profile/OneDrive
    $safeRoot    = 'C:\Temp\JTLS'
    $safeWork    = Join-Path $safeRoot 'work'
    $safeProfile = Join-Path $safeRoot 'chromium_profile'
    foreach ($p in @($safeRoot,$safeWork,$safeProfile)) {
        if (-not (Test-Path $p)) { $null = New-Item -ItemType Directory -Force -Path $p }
    }

    # Copy HTML to deterministic path
    $safeHtml = Join-Path $safeWork 'report.html'
    Copy-Item -Force $HtmlPath $safeHtml
    $htmlUrl = ([Uri]$safeHtml).AbsoluteUri

    $chromiumFlags = @(
        "--headless",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-extensions",
        "--no-first-run",
        "--no-default-browser-check",
        "--allow-file-access-from-files",
        "--virtual-time-budget=10000",
        "--run-all-compositor-stages-before-draw",
        "--print-to-pdf-no-header",
        "--kiosk-printing",
        "--print-to-pdf=""$pdfFull""",
        """$htmlUrl""",
        "--user-data-dir=""$safeProfile"""
    ) -join ' '

    # Edge
    $edge = Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'
    if (-not (Test-Path $edge)) { $edge = Join-Path ${env:ProgramFiles} 'Microsoft\Edge\Application\msedge.exe' }
    if (Invoke-JTLSBrowser -Exe $edge  -Flags $chromiumFlags -ExpectedPdf $pdfFull -TimeoutSeconds 60) { return $pdfFull }

    # Chrome
    $chrome = Join-Path ${env:ProgramFiles(x86)} 'Google\Chrome\Application\chrome.exe'
    if (-not (Test-Path $chrome)) { $chrome = Join-Path ${env:ProgramFiles} 'Google\Chrome\Application\chrome.exe' }
    if (Invoke-JTLSBrowser -Exe $chrome -Flags $chromiumFlags -ExpectedPdf $pdfFull -TimeoutSeconds 60) { return $pdfFull }

    # wkhtmltopdf (optional)
    $wk = Join-Path ${env:ProgramFiles} 'wkhtmltopdf\bin\wkhtmltopdf.exe'
    if (-not (Test-Path $wk)) { $wk = Join-Path ${env:ProgramFiles(x86)} 'wkhtmltopdf\bin\wkhtmltopdf.exe' }
    if ($wk -and (Test-Path $wk)) {
        $args = "--enable-local-file-access `"$safeHtml`" `"$pdfFull`""
        if (Invoke-JTLSCli -Exe $wk-Arguments $args -ExpectedPdf $pdfFull -TimeoutSeconds 60) { return $pdfFull }
    }

    # LibreOffice (optional)
    $soffice = Join-Path ${env:ProgramFiles} 'LibreOffice\program\soffice.exe'
    if (-not (Test-Path $soffice)) { $soffice = Join-Path ${env:ProgramFiles(x86)} 'LibreOffice\program\soffice.exe' }
    if ($soffice -and (Test-Path $soffice)) {
        $outDir = [IO.Path]::GetDirectoryName($pdfFull)
        $args = "--headless --convert-to pdf --outdir `"$outDir`" `"$safeHtml`""
        if (Invoke-JTLSCli -Exe $soffice -Arguments $args -ExpectedPdf $pdfFull -TimeoutSeconds 60) { return $pdfFull } {
            $guess = Join-Path $outDir (([IO.Path]::GetFileNameWithoutExtension($safeHtml)) + '.pdf')
            if ($guess -ne $pdfFull -and (Test-Path $guess)) { Move-Item -Force $guess $pdfFull }
            return $pdfFull
        }
    }

    throw 'No HTML→PDF method succeeded. Tried Edge, Chrome, wkhtmltopdf, LibreOffice.'
}
#endregion

#region Email HTML (template-driven)
function New-LogisticsEmailHTML {
    <#
    .SYNOPSIS
        Generates the HTML email body using a tokenized template and/or data.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$TemplatePath = $null,
        [Parameter(Mandatory)][array]$LotData,
        [Parameter(Mandatory = $false)][string]$PDFPath,
        [Parameter(Mandatory = $false)][object]$Config = $null
    )
    try {
        $pickupAddress = if ($Config -and $Config.auction_info -and $Config.auction_info.pickup_address) { $Config.auction_info.pickup_address } elseif ($LotData -and $LotData.Count -gt 0 -and $LotData[0].Address) { $LotData[0].Address.Trim() } else { "[Pickup address not specified]" }
        $deliveryAddress = if ($Config -and $Config.delivery_address) { $Config.delivery_address } else { "1218 Lake Avenue, Ashtabula, OH 44004" }
        $logisticsContactName  = if ($Config -and $Config.auction_info -and $Config.auction_info.logistics_contact -and $Config.auction_info.logistics_contact.name) { $Config.auction_info.logistics_contact.name } else { "[Not available]" }
        $logisticsContactPhone = if ($Config -and $Config.auction_info -and $Config.auction_info.logistics_contact -and $Config.auction_info.logistics_contact.phone) { $Config.auction_info.logistics_contact.phone } else { "[Not available]" }
        $logisticsContactEmail = if ($Config -and $Config.auction_info -and $Config.auction_info.logistics_contact -and $Config.auction_info.logistics_contact.email) { $Config.auction_info.logistics_contact.email } else { "[Not available]" }
        $pickupDateTime   = if ($Config -and $Config.auction_info -and $Config.auction_info.pickup_datetime) { $Config.auction_info.pickup_datetime } else { "[To be determined]" }
        $deliveryDateTime = if ($Config -and $Config.auction_info -and $Config.auction_info.delivery_datetime) { $Config.auction_info.delivery_datetime } else { "[To be determined]" }
        $pickupRequirements   = if ($Config -and $Config.auction_info -and $Config.auction_info.pickup_requirements) { $Config.auction_info.pickup_requirements } else { "" }
        $deliveryRequirements = if ($Config -and $Config.auction_info -and $Config.auction_info.delivery_requirements) { $Config.auction_info.delivery_requirements } else { "" }
        $deliveryNotes = if ($Config -and $Config.auction_info -and $Config.auction_info.delivery_notes) { $Config.auction_info.delivery_notes } elseif ($Config -and $Config.auction_info -and $Config.auction_info.delivery_notice) { $Config.auction_info.delivery_notice } else { "" }

        if (-not $TemplatePath -and $Config -and $Config.template_path) { $TemplatePath = $Config.template_path }
        if (-not $TemplatePath) { $TemplatePath = Join-Path $PSScriptRoot "Templates\LogisticsEmail_Template_New.html" }
        if (-not (Test-Path $TemplatePath)) { throw ("Template not found at: {0}" -f $TemplatePath) }

        $template = Get-Content -Raw -Path $TemplatePath

        $itemLines = @()
        foreach ($lot in $LotData) {
            $num = $null
            foreach ($cand in @($lot.Lot, $lot.LotNumber, $lot.Number, $lot.'Lot #')) {
                if ($null -ne $cand -and "$cand".Trim() -ne "") { $num = "$cand".Trim(); break }
            }
            if (-not $num) { $num = "[n/a]" }

            $desc = $null
            foreach ($cand in @($lot.Description, $lot.Title, $lot.ItemDescription, $lot.Name)) {
                if ($null -ne $cand -and "$cand".Trim() -ne "") { $desc = "$cand".Trim(); break }
            }
            if (-not $desc) { $desc = "[no description]" }

            $qty = $null
            foreach ($cand in @($lot.Quantity, $lot.Qty, $lot.Amount)) {
                if ($null -ne $cand -and "$cand".ToString().Trim() -ne "") { $qty = "$cand".ToString().Trim(); break }
            }
            $qtyText = if ($qty) { " (Qty: $qty)" } else { "" }
            $itemLines += "<li><strong>Lot #${num}:</strong> $desc$qtyText</li>"
        }
        $itemsHtml = ($itemLines -join "`n")

        $val = { param($s) if ($null -eq $s -or "$s" -eq "") { "" } else { "$s" } }
        $html = $template
        $replacements = @{
            "{{LOT_COUNT}}"             = "$($LotData.Count)"
            "{{PICKUP_ADDRESS}}"        = & $val $pickupAddress
            "{{DELIVERY_ADDRESS}}"      = & $val $deliveryAddress
            "{{CONTACT_NAME}}"          = & $val $logisticsContactName
            "{{CONTACT_PHONE}}"         = & $val $logisticsContactPhone
            "{{CONTACT_EMAIL}}"         = & $val $logisticsContactEmail
            "{{PICKUP_DATETIME}}"       = & $val $pickupDateTime
            "{{PICKUP_REQUIREMENTS}}"   = & $val $pickupRequirements
            "{{DELIVERY_DATETIME}}"     = & $val $deliveryDateTime
            "{{DELIVERY_REQUIREMENTS}}" = & $val $deliveryRequirements
            "{{DELIVERY_NOTES}}"        = & $val $deliveryNotes
            "{{ITEMS_HTML}}"            = $itemsHtml
        }
        foreach ($k in $replacements.Keys) { $html = $html.Replace($k, [string]$replacements[$k]) }
        return $html
    }
    catch {
        Write-JTLSLog -Message "Failed to generate HTML email: $_" -Level ERROR
        throw
    }
}
#endregion

#region Outlook
function New-OutlookDraftEmail {
    <#
    .SYNOPSIS
        Creates a draft email in Outlook with HTML body and attachments.
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)][string]$HTMLContent,
        [Parameter(Mandatory = $false)][string]$Subject = ("Freight Quote Request - {0}" -f (Get-Date -Format 'yyyy-MM-dd')),
        [Parameter(Mandatory = $false)][string[]]$Attachments = @(),
        [Parameter(Mandatory = $false)][string]$To = "",
        [Parameter(Mandatory = $false)][switch]$Display
    )

    try {
        Write-JTLSLog -Message 'Creating Outlook draft email' -Level INFO
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)
        $mail.Subject = $Subject
        $mail.HTMLBody = $HTMLContent
        if ($To) { $mail.To = $To }

        foreach ($attachment in $Attachments) {
            if (Test-Path $attachment) {
                $abs = (Resolve-Path $attachment).Path
                $null = $mail.Attachments.Add($abs)
                Write-JTLSLog -Message ("Added attachment: {0}" -f (Split-Path $abs -Leaf)) -Level DEBUG
            }
            else {
                Write-JTLSLog -Message ("Attachment not found: {0}" -f $attachment) -Level WARNING
            }
        }
        $mail.Save()
        if ($Display) { $mail.Display() }

        Write-JTLSLog -Message 'Outlook draft email created successfully' -Level SUCCESS

        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($mail)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($outlook)
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        return $true
    }
    catch {
        Write-JTLSLog -Message "Failed to create Outlook email: $_" -Level ERROR
        Write-JTLSLog -Message 'Make sure Outlook is installed and configured' -Level WARNING
        return $false
    }
}
#endregion

#region Dashboard / Report
function Show-Dashboard {
    <#
    .SYNOPSIS
        Displays an interactive processing dashboard.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Stats)

    Write-Host "`nPress Enter to view dashboard..." -ForegroundColor Cyan
    $null = Read-Host

    $endTime = Get-Date
    $duration = $endTime - $Stats.StartTime

    Clear-Host
    Write-Host ""
    Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║         LOGISTICS EMAIL AUTOMATION DASHBOARD              ║" -ForegroundColor Cyan
    Write-Host "╠════════════════════════════════════════════════════════════╣" -ForegroundColor Cyan

    $lotsWithImages = if ($Stats.TotalLots -gt 0 -and $Stats.ImagesMissing -ge 0) { $Stats.TotalLots - $Stats.ImagesMissing } else { $Stats.TotalLots }
    $successRate = if ($Stats.TotalLots -gt 0) { [math]::Round(($lotsWithImages / $Stats.TotalLots) * 100, 2) } else { 0 }
    $successRate = [math]::Min($successRate, 100)

    $barLength = 40
    $filled = [math]::Floor($successRate / 100 * $barLength)
    $filled = [math]::Max(0, [math]::Min($filled, $barLength))
    $empty  = $barLength - $filled
    $bar = ("█" * $filled) + ("░" * $empty)

    Write-Host ("Total Lots Processed: {0}" -f $Stats.TotalLots)
    Write-Host ("Images Found:         {0}" -f $Stats.ImagesFound)
    Write-Host ("Images Missing:       {0}" -f $Stats.ImagesMissing)
    Write-Host ("PDFs Generated:       {0}" -f $Stats.PDFsGenerated)
    Write-Host ("Emails Generated:     {0}" -f $Stats.EmailsGenerated)
    Write-Host ("Processing Time:      {0:mm\:ss}" -f $duration)
    Write-Host ("Coverage:             [{0}] {1}%" -f $bar, $successRate)

    Write-Host ""
    Write-Host "✅ Processing Complete!" -ForegroundColor Green
    Write-Host ("Logs saved to: {0}" -f $script:LogFile) -ForegroundColor Cyan
    Write-Host ""
}

function Export-ProcessingReport {
    <#
    .SYNOPSIS
        Exports an HTML report summarizing processing results.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][hashtable]$Stats,
        [Parameter(Mandatory)][array]$LotData
    )

    $reportPath = Join-Path $OutputPath ("ProcessingReport_{0}.html" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

    $reportHtml = @"
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

    foreach ($lot in $LotData) {
        $imageStatus = if (Test-Path (Join-Path $ImageDirectory ("{0}.jpg" -f $lot.Lot))) { "✅ Found" } else { "❌ Missing" }
        $reportHtml += @"
        <tr>
            <td>$($lot.Lot)</td>
            <td>$($lot.Description)</td>
            <td>$($lot.Quantity)</td>
            <td>$imageStatus</td>
        </tr>
"@
    }

    $reportHtml += @"
    </table>
</body>
</html>
"@

    $reportHtml | Out-File -FilePath $reportPath -Encoding UTF8
    Write-JTLSLog -Message ("Processing report exported to: {0}" -f $reportPath) -Level SUCCESS
    return $reportPath
}
#endregion

#region Main
try {
    # Align default directories with config location if provided
    if ($ConfigPath -and (Test-Path $ConfigPath)) {
        $configDir = Split-Path $ConfigPath -Parent
        $LogDirectory = Join-Path $configDir 'Logs'

        if ($PSBoundParameters.ContainsKey('OutputDirectory') -eq $false) {
            $OutputDirectory = Join-Path $configDir 'Output'
        }

        Write-Host "Using directories based on config location:" -ForegroundColor Cyan
        Write-Host ("  Logs:   {0}" -f $LogDirectory) -ForegroundColor Gray
        Write-Host ("  Output: {0}" -f $OutputDirectory) -ForegroundColor Gray
    }

    Start-JTLSLogging -LogDir $LogDirectory

    if (-not (Test-Path $OutputDirectory)) {
        $null = New-Item -ItemType Directory -Path $OutputDirectory -Force
        Write-JTLSLog -Message ("Created output directory: {0}" -f $OutputDirectory) -Level INFO
    }

    # Load configuration
    $config = $null
    if ($ConfigPath) {
        $config = Get-AuctionConfig -ConfigFile $ConfigPath
        if ($config) { Write-JTLSLog -Message ("Using configuration from: {0}" -f $ConfigPath) -Level SUCCESS }
    }

    # Import data
    Write-JTLSLog -Message 'Starting data import process' -Level INFO
    $auctionData = Import-AuctionData -Path $CSVPath

    # Process images
    Write-JTLSLog -Message ("Processing lot images (max {0} per lot)" -f $MaxImagesPerLot) -Level INFO
    $lotImages = Get-LotImage -Lots $auctionData -ImageDir $ImageDirectory -MaxImagesPerLot $MaxImagesPerLot

    # Generate HTML gallery report
    $pdfFilePath = $null
    $imageReportPath = $null

    if ($lotImages.Count -gt 0) {
        Write-JTLSLog -Message 'Generating HTML image report' -Level INFO
        $imageReportPath = New-LotHtmlReport -Images $lotImages -OutputPath $OutputDirectory -AuctionName 'Auction'

        # Convert HTML → PDF using the built-in fallback chain
        $pdfTarget = Join-Path $OutputDirectory ("AuctionLots_{0}.pdf" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
        try {
            $pdfFilePath = Convert-ToPdf -HtmlPath $imageReportPath -PdfPath $pdfTarget
            if ($pdfFilePath -and (Test-Path $pdfFilePath)) {
                Write-JTLSLog -Message ("PDF created successfully: {0}" -f (Split-Path $pdfFilePath -Leaf)) -Level SUCCESS
                $folder = Split-Path $pdfFilePath -Parent
                Start-Process explorer.exe -ArgumentList ("/select,`"{0}`"" -f $pdfFilePath)
            }
            else {
                Write-JTLSLog -Message 'PDF conversion returned no file, falling back to HTML report' -Level WARNING
                $pdfFilePath = $imageReportPath
            }
        }
        catch {
            Write-JTLSLog -Message "Error during PDF conversion: $_" -Level ERROR
            Write-JTLSLog -Message 'Falling back to HTML report' -Level WARNING
            $pdfFilePath = $imageReportPath
        }
    }
    else {
        Write-JTLSLog -Message 'No images found, skipping PDF generation' -Level WARNING
    }

    # Generate email HTML
    Write-JTLSLog -Message 'Generating HTML email with config data' -Level INFO
    $emailHtml = New-LogisticsEmailHTML -TemplatePath $TemplatePath -LotData $auctionData -PDFPath $pdfFilePath -Config $config

    # Determine subject
    $emailSubject = if ($config -and $config.email_subject) { $config.email_subject } else { "Freight Quote Request - $(Get-Date -Format 'yyyy-MM-dd')" }

    # Persist email HTML to file
    $outputDirAbsolute = if (Test-Path $OutputDirectory) { (Resolve-Path $OutputDirectory).Path } else { (New-Item -ItemType Directory -Path $OutputDirectory -Force).FullName }
    $htmlFilePath = Join-Path $outputDirAbsolute ("LogisticsEmail_{0}.html" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
    $emailHtml | Out-File -FilePath $htmlFilePath -Encoding UTF8
    Write-JTLSLog -Message ("HTML email saved to: {0}" -f $htmlFilePath) -Level SUCCESS
    Write-Host "`n================================================" -ForegroundColor Cyan
    Write-Host "EMAIL FILE LOCATION:" -ForegroundColor Yellow
    Write-Host $htmlFilePath -ForegroundColor Green
    Write-Host "================================================`n" -ForegroundColor Cyan

    # Attachments
    $attachmentList = @()
    if ($pdfFilePath) { $attachmentList += $pdfFilePath }

    # Outlook
    if ($CreateOutlookDraft) {
        $null = New-OutlookDraftEmail -HTMLContent $emailHtml -Subject $emailSubject -Attachments $attachmentList -Display
    }
    else {
        Write-JTLSLog -Message 'Outlook draft creation skipped (use -CreateOutlookDraft to enable)' -Level INFO
        Start-Process $htmlFilePath
    }

    # Report
    try {
        $reportPath = Export-ProcessingReport -OutputPath $OutputDirectory -Stats $script:ProcessingStats -LotData $auctionData
    }
    catch {
        Write-JTLSLog -Message ("Skipping processing report: {0}" -f $_.Exception.Message) -Level WARNING
    }

    $script:ProcessingStats.ProcessedLots = $auctionData.Count

    if ($ShowDashboard) {
        Show-Dashboard -Stats $script:ProcessingStats
    }

    Write-JTLSLog -Message '=== Logistics Email Automation Completed Successfully ===' -Level SUCCESS
    Write-JTLSLog -Message ("Output files saved to: {0}" -f $OutputDirectory) -Level INFO

    [pscustomobject]@{
        Success    = $true
        HTMLFile   = $htmlFilePath
        PDFFile    = $pdfFilePath
        ReportFile = $reportPath
        Statistics = $script:ProcessingStats
        LogFile    = $script:LogFile
    }
}
catch {
    Write-JTLSLog -Message ("Fatal error: {0}" -f $_) -Level ERROR
    Write-JTLSLog -Message ("Stack Trace: {0}" -f $_.ScriptStackTrace) -Level ERROR
    [pscustomobject]@{
        Success     = $false
        Error       = $_.Exception.Message
        LogFile     = $script:LogFile
        ErrorLogFile= $script:ErrorLogFile
    }
    throw
}
finally {
    Write-JTLSLog -Message 'Script execution completed' -Level INFO
    if ($script:ProcessingStats -and $script:ProcessingStats.StartTime) {
        Write-JTLSLog -Message ("Total execution time: {0}" -f ((Get-Date) - $script:ProcessingStats.StartTime)) -Level INFO
    }
}
#endregion
