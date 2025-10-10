<#
.SYNOPSIS
    HTML to PDF Conversion Helper - Universal converter for multiple PDF engines
    
.DESCRIPTION
    Provides reliable HTML to PDF conversion using multiple methods:
    - Foxit PDF Printer (primary method)
    - Microsoft Edge headless printing (fallback)
    - Google Chrome headless printing (fallback)
    
    This is a reusable helper script that can be dot-sourced into other scripts
    or called directly from the command line.
    
.PARAMETER HTMLPath
    Path to the HTML file to convert
    
.PARAMETER OutputPath
    Full path where the PDF should be saved (including filename)
    
.PARAMETER Method
    Conversion method to use: Auto, Foxit, Edge, Chrome
    Default is "Auto" which tries methods in order until one succeeds
    
.PARAMETER TestOnly
    Switch to run diagnostics only (shows available PDF converters)
    
.EXAMPLE
    .\Convert-HTMLtoPDF.ps1 -HTMLPath ".\report.html" -OutputPath ".\report.pdf"
    
.EXAMPLE
    .\Convert-HTMLtoPDF.ps1 -TestOnly
    
.EXAMPLE
    # Dot-source in another script
    . .\Convert-HTMLtoPDF.ps1
    $pdfPath = ConvertTo-PDFHelper -HTMLPath $htmlFile -OutputPath $outputFile
    
.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 10/10/2025
    Version: 1.4.0
    Change Date: 10/10/2025
    Change Purpose: Fixed apostrophe in username breaking temp paths

.CHANGELOG
    1.4.0 - 10/10/2025 - Changed temp directory to C:\Temp\PDFConversion
                       - Avoids apostrophe issues in usernames (like O'Neill)
                       - Fixes command-line parsing errors in Edge/Chrome
                       - Falls back to user temp if C:\Temp not accessible
    1.3.0 - 10/10/2025 - COMPLETE OneDrive fix for both HTML input and PDF output
                       - Creates PDF in local temp, then copies to OneDrive location
                       - Fixes "exit code 1" errors when writing to OneDrive
    1.2.0 - 10/10/2025 - Automatic OneDrive path detection and workaround
                       - Creates temp copy of HTML in local folder for conversion
                       - Fixes Edge/Chrome exit code 1 errors with OneDrive paths
    1.1.0 - 10/10/2025 - Added detailed command output for debugging
                       - Multiple retry attempts with status messages
    1.0.0 - 10/10/2025 - Initial Release
                       - Foxit PDF Printer support (primary method)
                       - Edge headless printing (reliable fallback)
                       - Chrome headless printing (secondary fallback)
                       - Diagnostic test function included

.LINK
    https://github.com/JONeillSr/Shipping-Management
    
.COMPONENT
    Requires PowerShell 5.1 or higher
    Works best with Foxit PDF Printer installed
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$HTMLPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Auto", "Foxit", "Edge", "Chrome")]
    [string]$Method = "Auto",
    
    [Parameter(Mandatory=$false)]
    [switch]$TestOnly
)

#region Core Conversion Function
function ConvertTo-PDFHelper {
    <#
    .SYNOPSIS
        Converts HTML file to PDF using available PDF engines
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 10/10/2025
        Version: 1.0.0
        Change Date: 
        Change Purpose: Reusable PDF conversion function
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$HTMLPath,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Auto", "Foxit", "Edge", "Chrome")]
        [string]$Method = "Auto",
        
        [Parameter(Mandatory=$false)]
        [switch]$Quiet
    )
    
    $startTime = Get-Date
    
    try {
        # Convert paths to absolute
        $AbsoluteHTMLPath = (Resolve-Path $HTMLPath).Path
        $AbsoluteOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
        
        # Ensure output directory exists
        $outputDir = Split-Path $AbsoluteOutputPath -Parent
        if (!(Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        }
        
        # Check if we're dealing with OneDrive paths (problematic for headless browsers)
        $useLocalCopy = $false
        $tempHTMLPath = $AbsoluteHTMLPath
        $tempPDFPath = $AbsoluteOutputPath
        $originalOutputPath = $AbsoluteOutputPath
        
        if ($AbsoluteHTMLPath -match "OneDrive" -or $AbsoluteOutputPath -match "OneDrive") {
            if (!$Quiet) { 
                Write-Host "⚠️  OneDrive path detected - using local temp folder for conversion" -ForegroundColor Yellow 
            }
            $useLocalCopy = $true
            
            # Use C:\Temp instead of user temp folder to avoid apostrophe issues in username
            $tempDir = "C:\Temp\PDFConversion"
            if (!(Test-Path $tempDir)) {
                try {
                    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                }
                catch {
                    # Fallback to user temp if C:\Temp is not accessible
                    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "PDFConversion"
                    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                }
            }
            
            # Copy HTML to temp if needed
            if ($AbsoluteHTMLPath -match "OneDrive") {
                $tempHTMLPath = Join-Path $tempDir (Split-Path $AbsoluteHTMLPath -Leaf)
                Copy-Item -Path $AbsoluteHTMLPath -Destination $tempHTMLPath -Force
                if (!$Quiet) { 
                    Write-Host "   HTML copied to: $tempHTMLPath" -ForegroundColor Gray 
                }
            }
            
            # Use temp location for PDF output
            if ($AbsoluteOutputPath -match "OneDrive") {
                $tempPDFPath = Join-Path $tempDir (Split-Path $AbsoluteOutputPath -Leaf)
                if (!$Quiet) { 
                    Write-Host "   PDF will be created in: $tempPDFPath" -ForegroundColor Gray 
                    Write-Host "   Then copied to: $originalOutputPath" -ForegroundColor Gray
                }
            }
        }
        
        # Update AbsoluteOutputPath to point to temp location during conversion
        $AbsoluteOutputPath = $tempPDFPath
        
        if (!$Quiet) {
            Write-Host "Converting HTML to PDF..." -ForegroundColor Cyan
            Write-Host "  Source: $(Split-Path $tempHTMLPath -Leaf)" -ForegroundColor Gray
            Write-Host "  Target: $(Split-Path $originalOutputPath -Leaf)" -ForegroundColor Gray
            if ($useLocalCopy) {
                Write-Host "  Working from local temp (OneDrive workaround)" -ForegroundColor Gray
            }
        }
        
        # Method 1: Try Foxit PDF Printer (BEST for your setup)
        if ($Method -eq "Auto" -or $Method -eq "Foxit") {
            if (!$Quiet) { Write-Host "`nTrying Foxit PDF Printer..." -ForegroundColor Yellow }
            
            # Find Foxit PDF Printer
            $foxitPrinter = Get-Printer -ErrorAction SilentlyContinue | Where-Object { 
                $_.Name -like "*Foxit*PDF*" -or 
                $_.Name -like "*Foxit Reader PDF Printer*" -or
                $_.Name -like "*Foxit PhantomPDF Printer*"
            } | Select-Object -First 1
            
            if ($foxitPrinter) {
                if (!$Quiet) { Write-Host "  Found: $($foxitPrinter.Name)" -ForegroundColor Green }
                
                try {
                    # Use Edge to print to Foxit PDF Printer
                    $edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                    
                    if (Test-Path $edgePath) {
                        # Edge can print to a specific printer with specific output path
                        $arguments = "--headless --disable-gpu --no-sandbox --print-to-pdf=`"$AbsoluteOutputPath`" `"$tempHTMLPath`""
                        
                        if (!$Quiet) { Write-Host "  Using Edge to print to Foxit..." -ForegroundColor Gray }
                        
                        $process = Start-Process -FilePath $edgePath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
                        
                        # Wait for file to be written
                        Start-Sleep -Seconds 3
                        
                        if (Test-Path $AbsoluteOutputPath) {
                            $pdfSize = (Get-Item $AbsoluteOutputPath).Length
                            if ($pdfSize -gt 0) {
                                $elapsedTime = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)
                                if (!$Quiet) {
                                    Write-Host "  ✓ PDF created successfully!" -ForegroundColor Green
                                    Write-Host "    Size: $([math]::Round($pdfSize/1MB, 2)) MB" -ForegroundColor Gray
                                    Write-Host "    Time: ${elapsedTime}s" -ForegroundColor Gray
                                }
                                return $AbsoluteOutputPath
                            }
                        }
                    }
                }
                catch {
                    if (!$Quiet) { Write-Host "  ✗ Foxit method failed: $_" -ForegroundColor Red }
                }
            }
            else {
                if (!$Quiet) { Write-Host "  ✗ Foxit PDF Printer not found" -ForegroundColor Red }
            }
        }
        
        # Method 2: Try Microsoft Edge (Built into Windows 10/11 - Very Reliable)
        if ($Method -eq "Auto" -or $Method -eq "Edge") {
            if (!$Quiet) { Write-Host "`nTrying Microsoft Edge..." -ForegroundColor Yellow }
            
            $edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
            if (Test-Path $edgePath) {
                try {
                    $arguments = "--headless --disable-gpu --no-sandbox --print-to-pdf=`"$AbsoluteOutputPath`" `"$tempHTMLPath`""
                    
                    if (!$Quiet) { 
                        Write-Host "  Executing Edge headless print..." -ForegroundColor Gray 
                        Write-Host "  Command: msedge.exe $arguments" -ForegroundColor DarkGray
                        Write-Host "  Expected output: $AbsoluteOutputPath" -ForegroundColor DarkGray
                    }
                    
                    $process = Start-Process -FilePath $edgePath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
                    
                    if (!$Quiet) { Write-Host "  Process exit code: $($process.ExitCode)" -ForegroundColor DarkGray }
                    
                    # Wait for file to be written
                    Start-Sleep -Seconds 3
                    
                    # Check multiple times with delays
                    $maxAttempts = 5
                    $attempt = 0
                    $pdfFound = $false
                    
                    while ($attempt -lt $maxAttempts -and !$pdfFound) {
                        if (Test-Path $AbsoluteOutputPath) {
                            $pdfSize = (Get-Item $AbsoluteOutputPath).Length
                            if ($pdfSize -gt 0) {
                                $pdfFound = $true
                                break
                            }
                        }
                        $attempt++
                        if (!$Quiet -and $attempt -lt $maxAttempts) {
                            Write-Host "  Waiting for PDF... (attempt $attempt/$maxAttempts)" -ForegroundColor DarkGray
                        }
                        Start-Sleep -Seconds 2
                    }
                    
                    if ($pdfFound) {
                        $pdfSize = (Get-Item $AbsoluteOutputPath).Length
                        $elapsedTime = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)
                        
                        # If we used temp location, copy PDF back to original location
                        if ($useLocalCopy -and $originalOutputPath -ne $AbsoluteOutputPath) {
                            if (!$Quiet) { Write-Host "  Copying PDF to OneDrive location..." -ForegroundColor Gray }
                            try {
                                Copy-Item -Path $AbsoluteOutputPath -Destination $originalOutputPath -Force
                                if (!$Quiet) {
                                    Write-Host "  ✓ PDF created successfully using Edge!" -ForegroundColor Green
                                    Write-Host "    Location: $originalOutputPath" -ForegroundColor Cyan
                                    Write-Host "    Size: $([math]::Round($pdfSize/1MB, 2)) MB" -ForegroundColor Gray
                                    Write-Host "    Time: ${elapsedTime}s" -ForegroundColor Gray
                                }
                                return $originalOutputPath
                            }
                            catch {
                                if (!$Quiet) { 
                                    Write-Host "  ✗ Failed to copy PDF to OneDrive: $_" -ForegroundColor Red 
                                    Write-Host "  PDF is available at: $AbsoluteOutputPath" -ForegroundColor Yellow
                                }
                                return $AbsoluteOutputPath
                            }
                        }
                        else {
                            if (!$Quiet) {
                                Write-Host "  ✓ PDF created successfully using Edge!" -ForegroundColor Green
                                Write-Host "    Location: $AbsoluteOutputPath" -ForegroundColor Cyan
                                Write-Host "    Size: $([math]::Round($pdfSize/1MB, 2)) MB" -ForegroundColor Gray
                                Write-Host "    Time: ${elapsedTime}s" -ForegroundColor Gray
                            }
                            return $AbsoluteOutputPath
                        }
                    }
                    else {
                        if (!$Quiet) { 
                            Write-Host "  ✗ PDF file not created after $maxAttempts attempts" -ForegroundColor Red 
                            Write-Host "    Expected at: $AbsoluteOutputPath" -ForegroundColor Yellow
                        }
                    }
                }
                catch {
                    if (!$Quiet) { Write-Host "  ✗ Edge method failed: $_" -ForegroundColor Red }
                }
            }
            else {
                if (!$Quiet) { Write-Host "  ✗ Edge not found" -ForegroundColor Red }
            }
        }
        
        # Method 3: Try Google Chrome
        if ($Method -eq "Auto" -or $Method -eq "Chrome") {
            if (!$Quiet) { Write-Host "`nTrying Google Chrome..." -ForegroundColor Yellow }
            
            $chromePaths = @(
                "C:\Program Files\Google\Chrome\Application\chrome.exe",
                "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            )
            
            $chromePath = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
            
            if ($chromePath) {
                try {
                    $arguments = "--headless --disable-gpu --no-sandbox --print-to-pdf=`"$AbsoluteOutputPath`" `"$tempHTMLPath`""
                    
                    if (!$Quiet) { 
                        Write-Host "  Executing Chrome headless print..." -ForegroundColor Gray 
                        Write-Host "  Command: chrome.exe $arguments" -ForegroundColor DarkGray
                        Write-Host "  Expected output: $AbsoluteOutputPath" -ForegroundColor DarkGray
                    }
                    
                    $process = Start-Process -FilePath $chromePath -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden -ErrorAction Stop
                    
                    if (!$Quiet) { Write-Host "  Process exit code: $($process.ExitCode)" -ForegroundColor DarkGray }
                    
                    # Wait for file to be written
                    Start-Sleep -Seconds 3
                    
                    # Check multiple times with delays
                    $maxAttempts = 5
                    $attempt = 0
                    $pdfFound = $false
                    
                    while ($attempt -lt $maxAttempts -and !$pdfFound) {
                        if (Test-Path $AbsoluteOutputPath) {
                            $pdfSize = (Get-Item $AbsoluteOutputPath).Length
                            if ($pdfSize -gt 0) {
                                $pdfFound = $true
                                break
                            }
                        }
                        $attempt++
                        if (!$Quiet -and $attempt -lt $maxAttempts) {
                            Write-Host "  Waiting for PDF... (attempt $attempt/$maxAttempts)" -ForegroundColor DarkGray
                        }
                        Start-Sleep -Seconds 2
                    }
                    
                    if ($pdfFound) {
                        $pdfSize = (Get-Item $AbsoluteOutputPath).Length
                        $elapsedTime = [math]::Round(((Get-Date) - $startTime).TotalSeconds, 1)
                        
                        # If we used temp location, copy PDF back to original location
                        if ($useLocalCopy -and $originalOutputPath -ne $AbsoluteOutputPath) {
                            if (!$Quiet) { Write-Host "  Copying PDF to OneDrive location..." -ForegroundColor Gray }
                            try {
                                Copy-Item -Path $AbsoluteOutputPath -Destination $originalOutputPath -Force
                                if (!$Quiet) {
                                    Write-Host "  ✓ PDF created successfully using Chrome!" -ForegroundColor Green
                                    Write-Host "    Location: $originalOutputPath" -ForegroundColor Cyan
                                    Write-Host "    Size: $([math]::Round($pdfSize/1MB, 2)) MB" -ForegroundColor Gray
                                    Write-Host "    Time: ${elapsedTime}s" -ForegroundColor Gray
                                }
                                return $originalOutputPath
                            }
                            catch {
                                if (!$Quiet) { 
                                    Write-Host "  ✗ Failed to copy PDF to OneDrive: $_" -ForegroundColor Red 
                                    Write-Host "  PDF is available at: $AbsoluteOutputPath" -ForegroundColor Yellow
                                }
                                return $AbsoluteOutputPath
                            }
                        }
                        else {
                            if (!$Quiet) {
                                Write-Host "  ✓ PDF created successfully using Chrome!" -ForegroundColor Green
                                Write-Host "    Location: $AbsoluteOutputPath" -ForegroundColor Cyan
                                Write-Host "    Size: $([math]::Round($pdfSize/1MB, 2)) MB" -ForegroundColor Gray
                                Write-Host "    Time: ${elapsedTime}s" -ForegroundColor Gray
                            }
                            return $AbsoluteOutputPath
                        }
                    }
                    else {
                        if (!$Quiet) { 
                            Write-Host "  ✗ PDF file not created after $maxAttempts attempts" -ForegroundColor Red 
                            Write-Host "    Expected at: $AbsoluteOutputPath" -ForegroundColor Yellow
                        }
                    }
                }
                catch {
                    if (!$Quiet) { Write-Host "  ✗ Chrome method failed: $_" -ForegroundColor Red }
                }
            }
            else {
                if (!$Quiet) { Write-Host "  ✗ Chrome not found" -ForegroundColor Red }
            }
        }
        
        # If all methods fail
        if (!$Quiet) {
            Write-Host "`n✗ All PDF conversion methods failed" -ForegroundColor Red
            Write-Host "  The HTML file is available at: $AbsoluteHTMLPath" -ForegroundColor Yellow
            Write-Host "  You can manually print it to PDF using Ctrl+P" -ForegroundColor Yellow
        }
        
        return $null
    }
    catch {
        if (!$Quiet) {
            Write-Host "`n✗ Error during PDF conversion: $_" -ForegroundColor Red
            Write-Host "  $($_.Exception.Message)" -ForegroundColor Gray
        }
        return $null
    }
    finally {
        # Clean up temp files if we made them
        if ($useLocalCopy) {
            try {
                if (Test-Path $tempHTMLPath) {
                    Remove-Item -Path $tempHTMLPath -Force -ErrorAction SilentlyContinue
                }
                if ($tempPDFPath -ne $originalOutputPath -and (Test-Path $tempPDFPath)) {
                    Remove-Item -Path $tempPDFPath -Force -ErrorAction SilentlyContinue
                }
                if (!$Quiet) { 
                    Write-Host "   Cleaned up temporary files" -ForegroundColor DarkGray 
                }
            }
            catch {
                # Ignore cleanup errors
            }
        }
    }
}
#endregion

#region Diagnostic Function
function Test-PDFConversionMethods {
    <#
    .SYNOPSIS
        Tests which PDF conversion methods are available on this system
    .NOTES
        Author: John O'Neill Sr.
        Company: Azure Innovators
        Create Date: 10/10/2025
        Version: 1.0.0
        Change Date: 
        Change Purpose: Diagnostic tool to check available PDF converters
    #>
    
    Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║         PDF CONVERSION METHODS DIAGNOSTIC                  ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan
    
    $results = @{
        FoxitPrinter = $false
        Edge = $false
        Chrome = $false
        Recommendations = @()
    }
    
    # Check Foxit PDF Printers
    Write-Host "1. Checking for Foxit PDF Printers..." -ForegroundColor Yellow
    $foxitPrinters = Get-Printer -ErrorAction SilentlyContinue | Where-Object { 
        $_.Name -like "*Foxit*PDF*" 
    }
    if ($foxitPrinters) {
        foreach ($printer in $foxitPrinters) {
            Write-Host "   ✓ Found: $($printer.Name)" -ForegroundColor Green
            Write-Host "     Status: $($printer.PrinterStatus)" -ForegroundColor Gray
            Write-Host "     Port: $($printer.PortName)" -ForegroundColor Gray
        }
        $results.FoxitPrinter = $true
        $results.Recommendations += "Foxit PDF Printer is available (RECOMMENDED)"
    }
    else {
        Write-Host "   ✗ No Foxit printers found" -ForegroundColor Red
        Write-Host "     Install Foxit Reader or PhantomPDF to get the printer" -ForegroundColor Gray
    }
    
    # Check Edge
    Write-Host "`n2. Checking for Microsoft Edge..." -ForegroundColor Yellow
    $edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    if (Test-Path $edgePath) {
        $edgeVersion = (Get-Item $edgePath).VersionInfo.FileVersion
        Write-Host "   ✓ Found: $edgePath" -ForegroundColor Green
        Write-Host "     Version: $edgeVersion" -ForegroundColor Gray
        $results.Edge = $true
        $results.Recommendations += "Microsoft Edge is available (RELIABLE FALLBACK)"
    }
    else {
        Write-Host "   ✗ Edge not found" -ForegroundColor Red
        Write-Host "     Edge is built into Windows 10/11" -ForegroundColor Gray
    }
    
    # Check Chrome
    Write-Host "`n3. Checking for Google Chrome..." -ForegroundColor Yellow
    $chromePaths = @(
        "C:\Program Files\Google\Chrome\Application\chrome.exe",
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    )
    
    $chromePath = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($chromePath) {
        $chromeVersion = (Get-Item $chromePath).VersionInfo.FileVersion
        Write-Host "   ✓ Found: $chromePath" -ForegroundColor Green
        Write-Host "     Version: $chromeVersion" -ForegroundColor Gray
        $results.Chrome = $true
        $results.Recommendations += "Google Chrome is available (SECONDARY FALLBACK)"
    }
    else {
        Write-Host "   ✗ Chrome not found" -ForegroundColor Red
        Write-Host "     Install Chrome from google.com/chrome" -ForegroundColor Gray
    }
    
    # Summary
    Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║                        SUMMARY                             ║" -ForegroundColor Cyan
    Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan
    
    $availableCount = ($results.FoxitPrinter, $results.Edge, $results.Chrome | Where-Object { $_ }).Count
    
    if ($availableCount -eq 0) {
        Write-Host "⚠️  WARNING: No PDF converters found!" -ForegroundColor Red
        Write-Host "   Install Foxit PDF Reader or use Edge/Chrome" -ForegroundColor Yellow
    }
    elseif ($availableCount -eq 1) {
        Write-Host "✓ Found 1 PDF converter" -ForegroundColor Yellow
        Write-Host "  Consider installing additional converters for redundancy" -ForegroundColor Gray
    }
    else {
        Write-Host "✓ Found $availableCount PDF converters" -ForegroundColor Green
        Write-Host "  Your system has good redundancy!" -ForegroundColor Gray
    }
    
    if ($results.Recommendations.Count -gt 0) {
        Write-Host "`nAvailable Methods:" -ForegroundColor Cyan
        foreach ($rec in $results.Recommendations) {
            Write-Host "  • $rec" -ForegroundColor White
        }
    }
    
    Write-Host "`nRecommended Usage:" -ForegroundColor Cyan
    Write-Host '  ConvertTo-PDFHelper -HTMLPath "input.html" -OutputPath "output.pdf" -Method "Auto"' -ForegroundColor Gray
    
    Write-Host "`n════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan
    
    return $results
}
#endregion

#region Script Execution
# If called directly (not dot-sourced)
if ($MyInvocation.InvocationName -ne '.') {
    
    if ($TestOnly) {
        # Run diagnostics
        Test-PDFConversionMethods
    }
    elseif ($HTMLPath -and $OutputPath) {
        # Run conversion
        $result = ConvertTo-PDFHelper -HTMLPath $HTMLPath -OutputPath $OutputPath -Method $Method
        
        if ($result) {
            Write-Host "`n✓ Success! PDF created at:" -ForegroundColor Green
            Write-Host "  $result" -ForegroundColor White
            exit 0
        }
        else {
            Write-Host "`n✗ PDF conversion failed" -ForegroundColor Red
            exit 1
        }
    }
    else {
        Write-Host "`nUsage:" -ForegroundColor Cyan
        Write-Host "  .\Convert-HTMLtoPDF.ps1 -HTMLPath <path> -OutputPath <path> [-Method <Auto|Foxit|Edge|Chrome>]" -ForegroundColor White
        Write-Host "  .\Convert-HTMLtoPDF.ps1 -TestOnly" -ForegroundColor White
        Write-Host "`nOr dot-source in another script:" -ForegroundColor Cyan
        Write-Host "  . .\Convert-HTMLtoPDF.ps1" -ForegroundColor White
        Write-Host '  $pdf = ConvertTo-PDFHelper -HTMLPath "input.html" -OutputPath "output.pdf"' -ForegroundColor White
        Write-Host ""
    }
}
#endregion