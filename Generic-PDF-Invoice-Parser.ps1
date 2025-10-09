<#
.SYNOPSIS
    Generic PDF Invoice Parser for Multiple Auction Houses

.DESCRIPTION
    Advanced PDF parsing with support for multiple auction house formats, pattern learning,
    and intelligent data extraction. Works with Brolyn, Ritchie Bros, Purple Wave, GovDeals,
    and other auction platforms.

    v1.3.2 adds robust handling for invoices that print totals in multiple columns or where
    label/amount order is disrupted by PDF text reflow. The script now:
      - Captures Subtotal, Convenience Fee, Cash Total Due, and Credit Total Due using a
        small "window" after each label to avoid drifting into nearby columns.
      - Applies relationship logic: when Subtotal is known, CashTotal defaults to Subtotal;
        if a Convenience Fee is present, CreditTotal = Subtotal + ConvenienceFee.
      - Lets you choose the reported total via -PaymentMethod Cash|Credit (default Cash),
        or interactively with -PromptPayment.
      - Formats displayed totals with two decimals for readability.

.PARAMETER PDFPath
    Path to the PDF (or .txt extracted text) invoice to parse.

.PARAMETER OutputFormat
    Output format: JSON, CSV, Display, or Config (default: Display)
      - Display : prints a human-readable summary to the console
      - JSON    : writes the parsed result to InvoiceData_YYYYMMDD_HHMMSS.json
      - Config  : writes a logistics-oriented JSON config for freight quotes
      - CSV     : (reserved for future use)

.PARAMETER SavePattern
    Save extracted patterns for future parsing improvements (reserved for future use).

.PARAMETER DebugMode
    Enable debug output and save extracted text + parsed JSON to files in the current directory.

.PARAMETER GUI
    Launch a graphical UI (not yet implemented).

.PARAMETER PaymentMethod
    Which total to report when both are present: Cash or Credit (default: Cash).
    - Cash   -> Uses Cash Total Due (or Subtotal when CashTotal is not printed explicitly).
    - Credit -> Uses Credit Total Due (Subtotal + Convenience Fee when fee is present).

.PARAMETER PromptPayment
    Prompt at runtime to choose Cash or Credit (useful when you sometimes pay by card).
    If both totals can be determined, you’ll be asked which one to use; otherwise the script
    falls back to -PaymentMethod.

.EXAMPLE
    .\Generic-PDF-Invoice-Parser.ps1 -PDFPath ".\invoice.pdf"

    Parses the invoice and prints a formatted summary (Display) using the Cash total by default.

.EXAMPLE
    .\Generic-PDF-Invoice-Parser.ps1 -PDFPath ".\invoice.pdf" -PaymentMethod Credit

    Forces the script to report the Credit Total Due (e.g., Subtotal + Convenience Fee).

.EXAMPLE
    .\Generic-PDF-Invoice-Parser.ps1 -PDFPath ".\invoice.pdf" -PromptPayment

    Prompts you to choose Cash or Credit at runtime if both are available.

.EXAMPLE
    .\Generic-PDF-Invoice-Parser.ps1 -PDFPath ".\invoice.pdf" -OutputFormat JSON -DebugMode

    Saves the parsed data to JSON and also writes DEBUG_ExtractedText_*.txt and DEBUG_ParsedData_*.json.

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Create Date: 2025-01-08
    Version: 1.3.2
    Change Date: 2025-10-08
    Change Purpose:
      - Windowed label extraction to avoid cross-column drift
      - Relationship-based correction for Cash/Credit totals (cash=subtotal, credit=subtotal+fee)
      - Optional prompt for payment method selection
      - Fixed debug prints and consistent 2-decimal formatting in Display output
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$PDFPath,

    [Parameter(Mandatory=$false)]
    [ValidateSet("JSON", "CSV", "Display", "Config")]
    [string]$OutputFormat = "Display",

    [Parameter(Mandatory=$false)]
    [switch]$SavePattern,

    [Parameter(Mandatory=$false)]
    [switch]$DebugMode,

    [Parameter(Mandatory=$false)]
    [switch]$GUI,

    [Parameter(Mandatory=$false)]
    [ValidateSet("Cash","Credit")]
    [string]$PaymentMethod = "Cash",

    [Parameter(Mandatory=$false)]
    [switch]$PromptPayment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:PatternsFile = ".\Data\InvoicePatterns.json"
$script:LearnedPatterns = @()

#region PDF Extraction Functions
function Get-PDFTextContent {
    <#
    .SYNOPSIS
        Extracts text from PDF using multiple methods with fallback
    .NOTES
        Version: 1.3.0
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    Write-Output "`n📄 Extracting text from PDF..."
    Write-Output "   File: $(Split-Path $Path -Leaf)"

    # Method 1: Try pdftotext (xpdf-tools) if available
    $pdftotext = Get-Command pdftotext -ErrorAction SilentlyContinue
    if ($pdftotext) {
        try {
            $tempFile = [System.IO.Path]::GetTempFileName()
            & pdftotext -layout $Path $tempFile 2>&1 | Out-Null
            
            if (Test-Path $tempFile) {
                $extractedText = Get-Content $tempFile -Raw -Encoding UTF8
                Remove-Item $tempFile -Force
                
                if ($extractedText -and $extractedText.Length -gt 100) {
                    Write-Output "   ✅ Extracted $($extractedText.Length) characters using pdftotext"
                    return @{
                        Text = $extractedText
                        Method = "pdftotext"
                        Quality = "High"
                    }
                }
            }
        }
        catch {
            Write-Verbose "pdftotext extraction failed: $_"
        }
    }

    # Method 2: Try using .NET PDF libraries if available
    try {
        # Check if iTextSharp is available
        $iTextSharpPath = "$env:USERPROFILE\.nuget\packages\itextsharp\*\lib\net40\itextsharp.dll"
        $iTextDll = Get-ChildItem $iTextSharpPath -ErrorAction SilentlyContinue | Select-Object -First 1
        
        if ($iTextDll) {
            Add-Type -Path $iTextDll.FullName
            $reader = New-Object iTextSharp.text.pdf.PdfReader($Path)
            $extractedText = ""
            
            for ($page = 1; $page -le $reader.NumberOfPages; $page++) {
                $strategy = New-Object iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
                $currentText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $page, $strategy)
                $extractedText += $currentText + "`n"
            }
            
            $reader.Close()
            
            if ($extractedText.Length -gt 100) {
                Write-Output "   ✅ Extracted $($extractedText.Length) characters using iTextSharp"
                return @{
                    Text = $extractedText
                    Method = "iTextSharp"
                    Quality = "High"
                }
            }
        }
    }
    catch {
        Write-Verbose "iTextSharp extraction failed: $_"
    }

    # If we get here, no method worked
    Write-Warning "`n❌ Could not extract readable text from PDF"
    Write-Output "`n💡 This PDF requires a proper PDF parsing library."
    Write-Output "`nRECOMMENDED SOLUTIONS:"
    Write-Output ""
    Write-Output "Option 1 - Install xpdf-tools (EASIEST):"
    Write-Output "  1. Download from: https://www.xpdfreader.com/download.html"
    Write-Output "  2. Install xpdf-tools (includes pdftotext.exe)"
    Write-Output "  3. Add to PATH or place pdftotext.exe in Windows\System32"
    Write-Output "  4. Re-run this script"
    Write-Output ""
    Write-Output "Option 2 - Manual text extraction:"
    Write-Output "  1. Open the PDF in Adobe Reader"
    Write-Output "  2. File > Save As Other > Text"
    Write-Output "  3. Save as .txt file"
    Write-Output "  4. Use -PDFPath with the .txt file instead"
    Write-Output ""
    Write-Output "Option 3 - Use Adobe Reader:"
    Write-Output "  1. Open PDF in Adobe Reader"
    Write-Output "  2. Select All (Ctrl+A)"
    Write-Output "  3. Copy (Ctrl+C)"
    Write-Output "  4. Paste into a text file"
    Write-Output "  5. Save and use that file"
    
    return $null
}
#endregion

#region Pattern Recognition Functions
function Initialize-InvoicePattern {
    <#
    .SYNOPSIS
        Loads learned patterns from previous parses
    #>

    if (Test-Path $script:PatternsFile) {
        $script:LearnedPatterns = Get-Content $script:PatternsFile -Raw | ConvertFrom-Json
        Write-Verbose "Loaded $($script:LearnedPatterns.Count) learned patterns"
    }
    else {
        $script:LearnedPatterns = @()

        # Initialize with default patterns
        $script:LearnedPatterns += @{
            Vendor = "Brolyn Auctions"
            Patterns = @{
                CompanyIdentifier = 'Brolyn|BROLYN'
                Phone = '\(574\)\s*891-3111'
                Email = 'logistics@brolynauctions\.com'
                Address = '290\s+West\s+750\s+North.*?Howe.*?IN|1139\s+Haines.*?Sturgis.*?MI'
                PickupDates = '(?:load times|pickup).*?(\w+\s+\d{1,2}/\d{1,2})\s+thru\s+(\w+\s+\d{1,2}/\d{1,2})'
            }
        }

        $script:LearnedPatterns += @{
            Vendor = "Ritchie Bros"
            Patterns = @{
                CompanyIdentifier = 'Ritchie Bros|RITCHIE BROS|RB Auctions'
                Phone = '\(\d{3}\)\s*\d{3}-\d{4}'
                Email = '[a-zA-Z0-9._%+-]+@rbauction\.com|[a-zA-Z0-9._%+-]+@ritchiebros\.com'
                Address = '\d+.*?(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd).*?[A-Z]{2}\s+\d{5}'
                PickupDates = '(?:Removal|Pickup).*?(\d{1,2}/\d{1,2}/\d{4})'
            }
        }
    }
}

function Find-InvoiceVendor {
    <#
    .SYNOPSIS
        Identifies the auction vendor from invoice text with enhanced pattern matching
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$Text
    )

    Write-Output "`n🔍 Identifying vendor..."

    # Clean text for better matching
    $cleanText = $Text -replace '\s+', ' '

    # Look for Brolyn with more flexible matching
    if ($cleanText -match '(?i)brolyn|BROLYN') {
        Write-Output "   ✅ Detected: Brolyn Auctions"
        return $script:LearnedPatterns | Where-Object { $_.Vendor -eq "Brolyn Auctions" } | Select-Object -First 1
    }

    foreach ($pattern in $script:LearnedPatterns) {
        if ($cleanText -match $pattern.Patterns.CompanyIdentifier) {
            Write-Output "   ✅ Detected: $($pattern.Vendor)"
            return $pattern
        }
    }

    Write-Output "   ⚠️  Unknown vendor - using generic patterns"
    return @{
        Vendor = "Unknown"
        Patterns = @{
            CompanyIdentifier = ""
            Phone = '\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}'
            Email = '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
            Address = '\d+\s+[A-Za-z\s]+(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd|Drive|Dr|Lane|Ln)[,\s]+[A-Za-z\s]+[,\s]+[A-Z]{2}\s+\d{5}'
            PickupDates = '(\d{1,2}/\d{1,2}/\d{4})'
        }
    }
}

function Test-ValidPhoneNumber {
    <#
    .SYNOPSIS
        Validates if a phone number is legitimate (not a false positive)
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$Phone
    )

    # Extract just digits
    $digits = $Phone -replace '[^\d]', ''
    
    # Must be exactly 10 digits
    if ($digits.Length -ne 10) {
        return $false
    }

    # Area code (first 3 digits) must be valid (2-9 for first digit, 0-9 for others)
    $areaCode = [int]$digits.Substring(0, 3)
    if ($areaCode -lt 200 -or $areaCode -gt 999) {
        return $false
    }

    # Exchange (middle 3 digits) must be valid (2-9 for first digit)
    $exchange = [int]$digits.Substring(3, 3)
    if ($exchange -lt 200 -or $exchange -gt 999) {
        return $false
    }

    return $true
}

function Get-InvoiceData {
    <#
    .SYNOPSIS
        Extracts structured data from invoice text with enhanced pattern matching
    .DESCRIPTION
        Parses vendor, contacts, pickup info, items, and totals.
        Totals logic:
          - Uses a small character window after labels (e.g., "Cash Total Due:") to pull the
            nearest $amount and avoid column drift.
          - When Subtotal is present, sets CashTotal = Subtotal. If a Convenience Fee exists,
            sets CreditTotal = Subtotal + ConvenienceFee.
          - If -PromptPayment is set and both totals are available, asks user which to use.
          - Otherwise, selects based on -PaymentMethod (default Cash).
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$Text,

        [Parameter(Mandatory=$true)]
        [object]$VendorPattern,

        [Parameter(Mandatory=$true)]
        [ValidateSet("Cash","Credit")]
        [string]$PaymentMethod,

        [Parameter(Mandatory=$true)]
        [bool]$PromptPayment
    )

    Write-Output "`n📊 Extracting invoice data..."

    # Normalize whitespace while preserving line breaks for lot extraction
    $normalizedText = $Text -replace '[\r\n]+', ' '
    $normalizedText = $normalizedText -replace '\s+', ' '

    $data = @{
        Vendor = $VendorPattern.Vendor
        InvoiceNumber = $null
        InvoiceDate = $null
        ContactInfo = @{
            Phone = @()
            Email = @()
        }
        PickupAddresses = @()
        PickupDates = @()
        Items = @()
        Totals = @{
            Subtotal = $null
            Tax = $null
            Premium = $null
            Total = $null
            CashTotal = $null
            CreditTotal = $null
            ConvenienceFee = $null
        }
        SpecialNotes = @()
    }

    # Extract invoice number
    if ($normalizedText -match '(?:Invoice\s*#?\s*:?\s*)?(\d{4}-\d{6}-\d+)') {
        $data.InvoiceNumber = $Matches[1]
        Write-Output "   📋 Invoice #: $($data.InvoiceNumber)"
    }
    elseif ($normalizedText -match 'Invoice\s*#?\s*:?\s*([A-Z0-9-]+)') {
        $data.InvoiceNumber = $Matches[1]
        Write-Output "   📋 Invoice #: $($data.InvoiceNumber)"
    }

    # Extract invoice date
    if ($normalizedText -match 'Date:\s*(\d{1,2}/\d{1,2}/\d{4})') {
        $data.InvoiceDate = $Matches[1]
        Write-Output "   📅 Date: $($data.InvoiceDate)"
    }
    elseif ($normalizedText -match 'Invoice\s+Date\s*:?\s*(\d{1,2}[-/]\w+[-/]\d{4}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})') {
        $data.InvoiceDate = $Matches[1]
        Write-Output "   📅 Date: $($data.InvoiceDate)"
    }
    elseif ($normalizedText -match '\d{2}-\w{3}-\d{4}\s+\d{2}:\d{2}') {
        $data.InvoiceDate = $Matches[0]
        Write-Output "   📅 Date: $($data.InvoiceDate)"
    }

    # Phones
    $phonePattern = '\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}'
    $phoneMatches = [regex]::Matches($normalizedText, $phonePattern)
    foreach ($match in $phoneMatches) {
        $phone = $match.Value.Trim()
        if (Test-ValidPhoneNumber -Phone $phone) {
            $digits = $phone -replace '[^\d]', ''
            $formattedPhone = "($($digits.Substring(0,3))) $($digits.Substring(3,3))-$($digits.Substring(6,4))"
            if ($data.ContactInfo.Phone -notcontains $formattedPhone) {
                $data.ContactInfo.Phone += $formattedPhone
            }
        }
    }
    if ($data.ContactInfo.Phone.Count -gt 0) {
        Write-Output "   📞 Found $($data.ContactInfo.Phone.Count) phone number(s)"
    }

    # Emails
    $emailPattern = '\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b'
    $emailMatches = [regex]::Matches($normalizedText, $emailPattern)
    foreach ($match in $emailMatches) {
        $email = $match.Value.Trim().ToLower()
        if ($data.ContactInfo.Email -notcontains $email) {
            $data.ContactInfo.Email += $email
        }
    }
    if ($data.ContactInfo.Email.Count -gt 0) {
        Write-Output "   📧 Found $($data.ContactInfo.Email.Count) email(s)"
    }

    # Addresses (use original text with line breaks)
    $addressPatterns = @(
        '(\d+\s+(?:West|East|North|South)\s+\d+\s+(?:West|East|North|South)[^\n,]+?(?:Howe|Sturgis)[^\n,]+?(?:IN|MI)\s+\d{5})',
        '(\d+\s+[A-Za-z]+\s+(?:Blvd|Boulevard|Street|St|Avenue|Ave|Road|Rd|Drive|Dr)[^\n,]+?(?:Howe|Sturgis)[^\n,]+?(?:IN|MI)\s+\d{5})'
    )

    foreach ($pattern in $addressPatterns) {
        $addressMatches = [regex]::Matches($Text, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($match in $addressMatches) {
            $address = $match.Groups[1].Value.Trim()
            $address = $address -replace '\s+', ' '
            $address = $address -replace '\s*,\s*', ', '
            if ($address.Length -gt 15 -and $data.PickupAddresses -notcontains $address) {
                $data.PickupAddresses += $address
            }
        }
    }
    if ($data.PickupAddresses.Count -gt 0) {
        Write-Output "   📍 Found $($data.PickupAddresses.Count) address(es)"
    }

    # Pickup dates
    $uniqueDates = @{}
    if ($normalizedText -match '(?i)load\s+times\s+for\s+materials[^:]+:\s*((?:Monday|Tuesday|Wednesday|Thursday|Friday)\s+\d{1,2}/\d{1,2}(?:/\d{2,4})?\s+thru\s+(?:Monday|Tuesday|Wednesday|Thursday|Friday)\s+\d{1,2}/\d{1,2})') {
        $uniqueDates[$Matches[1].Trim()] = $true
    }
    if ($normalizedText -match '(?i)load\s+times\s+for\s+racking[^:]+:\s*((?:Monday|Tuesday|Wednesday|Thursday|Friday)\s+\d{1,2}/\d{1,2}(?:/\d{2,4})?\s+thru\s+(?:Monday|Tuesday|Wednesday|Thursday|Friday)\s+\d{1,2}/\d{1,2})') {
        $uniqueDates[$Matches[1].Trim()] = $true
    }
    if ($normalizedText -match '(?i)payment\s+must\s+be\s+received\s+by\s+((?:Monday|Tuesday|Wednesday|Thursday|Friday)\s+\d{1,2}/\d{1,2}/\d{2}\s+at\s+\d{1,2}[ap]m)') {
        $uniqueDates[$Matches[1].Trim()] = $true
    }
    $data.PickupDates = @($uniqueDates.Keys)
    if ($data.PickupDates.Count -gt 0) {
        Write-Output "   📅 Found $($data.PickupDates.Count) pickup date range(s)"
    }

    # Items (lot extraction)
    $lotLines = $Text -split "`n"
    $currentLot = $null
    foreach ($lineRaw in $lotLines) {
        $line = $lineRaw.Trim()
        if ($line -match '^(\d{2,5})\s+\d{4}\s+(.+)') {
            $lotNum = $Matches[1]
            $desc = $Matches[2].Trim()
            if ($desc -notmatch '(?i)^(?:Invoice|CR\s+\d+|Location:|Page\s+\d+|Date:)' -and 
                $desc.Length -ge 15 -and $desc.Length -le 500) {
                $desc = $desc -replace '\s+', ' '
                $desc = $desc -replace '^\s*-\s*', ''
                $existingLot = $data.Items | Where-Object { $_.LotNumber -eq $lotNum }
                if (-not $existingLot) {
                    $data.Items += @{
                        LotNumber = $lotNum
                        Description = $desc
                    }
                    $currentLot = @{
                        LotNumber = $lotNum
                        Description = $desc
                    }
                }
            }
        }
        elseif ($currentLot -and $line -match '^[A-Z]' -and $line.Length -gt 10 -and $line.Length -lt 200) {
            if ($line -notmatch '^\d+\s+\d{4}' -and $line -notmatch '^Location:') {
                $idx = $data.Items.Count - 1
                if ($idx -ge 0 -and $data.Items[$idx].LotNumber -eq $currentLot.LotNumber) {
                    $data.Items[$idx].Description += " " + $line.Trim()
                }
            }
        }
    }
    if ($data.Items.Count -gt 0) {
        Write-Output "   📦 Found $($data.Items.Count) lot item(s)"
    }

    # ====== ROBUST: Financial totals (windowed after labels + relationship rules) ======
    $norm = ($Text -replace '\s+', ' ').Trim()

    function Get-ParsedAmount([string]$s) { if ($s) { [decimal]($s -replace ',','') } else { $null } }

    function GetAmountAfterLabel([string]$text, [string]$label, [int]$window = 100) {
        $m = [regex]::Match($text, [regex]::Escape($label), 'IgnoreCase')
        if (-not $m.Success) { return $null }
        $start = [Math]::Min($text.Length, $m.Index + $m.Length)
        $len   = [Math]::Min($window, [Math]::Max(0, $text.Length - $start))
        if ($len -le 0) { return $null }
        $slice = $text.Substring($start, $len)
        $m2 = [regex]::Match($slice, '\$\s*([\d,]+\.\d{2})')
        if ($m2.Success) { return Get-ParsedAmount $m2.Groups[1].Value }
        return $null
    }

    # Windowed captures (may be imperfect on column-wrap PDFs)
    $capturedSubtotal       = GetAmountAfterLabel $norm 'SubTotal:'           80
    $capturedCashTotal      = GetAmountAfterLabel $norm 'Cash Total Due:'     80
    $capturedConvenienceFee = GetAmountAfterLabel $norm 'Convenience Fee'     80
    $capturedCreditTotal    = GetAmountAfterLabel $norm 'Credit Total Due:'   80
    $capturedGrandTotal     = GetAmountAfterLabel $norm 'Grand Total:'        80

    # Assign what we captured
    if ($capturedSubtotal)       { $data.Totals.Subtotal       = $capturedSubtotal }
    if ($capturedConvenienceFee) { $data.Totals.ConvenienceFee = $capturedConvenienceFee }
    if ($capturedCashTotal)      { $data.Totals.CashTotal      = $capturedCashTotal }
    if ($capturedCreditTotal)    { $data.Totals.CreditTotal    = $capturedCreditTotal }

    # Relationship rules (deterministic behavior)
    if ($data.Totals.Subtotal) {
        $sub = [decimal]$data.Totals.Subtotal
        $fee = if ($data.Totals.ConvenienceFee) { [decimal]$data.Totals.ConvenienceFee } else { $null }

        # Cash = Subtotal (Brolyn: usually no added tax/premium in sample)
        if (-not $data.Totals.CashTotal -or ($data.Totals.CashTotal -lt $sub)) {
            $data.Totals.CashTotal = $sub
        }
        # If Cash accidentally equals the fee (common drift issue), fix it
        if ($fee -and $data.Totals.CashTotal -eq $fee) {
            $data.Totals.CashTotal = $sub
        }

        # Credit = Subtotal + fee (when fee known)
        if ($fee) {
            $data.Totals.CreditTotal = $sub + $fee
        } elseif (-not $data.Totals.CreditTotal -and $capturedGrandTotal) {
            # If only a generic grand total exists, treat it as cash-like (no fee)
            $data.Totals.CreditTotal = $null
            if (-not $data.Totals.CashTotal) { $data.Totals.CashTotal = $capturedGrandTotal }
        }
    }

    # Choose payment method
    $selectedMethod = $PaymentMethod
    if ($PromptPayment -and $data.Totals.CashTotal -and ($data.Totals.CreditTotal -or $data.Totals.ConvenienceFee)) {
        $answer = Read-Host "Payment method for totals (Cash/Credit) [$selectedMethod]"
        if ($answer -match '^(?i)credit$') { $selectedMethod = 'Credit' }
        elseif ($answer -match '^(?i)cash$') { $selectedMethod = 'Cash' }
    }

    switch ($selectedMethod) {
        'Credit' {
            if ($data.Totals.CreditTotal) {
                $data.Totals.Total = $data.Totals.CreditTotal
            }
            elseif ($data.Totals.Subtotal -and $data.Totals.ConvenienceFee) {
                $data.Totals.Total = [decimal]$data.Totals.Subtotal + [decimal]$data.Totals.ConvenienceFee
            }
            else {
                $data.Totals.Total = $data.Totals.CashTotal
            }
        }
        default {
            $data.Totals.Total = $data.Totals.CashTotal
        }
    }

    # Final sanity
    if ($data.Totals.Total -and $data.Totals.Subtotal -and $data.Totals.Total -lt $data.Totals.Subtotal) {
        Write-Warning "Total ($($data.Totals.Total)) < Subtotal ($($data.Totals.Subtotal)); correcting to Subtotal for Cash."
        $data.Totals.Total = $data.Totals.Subtotal
        $data.Totals.CashTotal = $data.Totals.Subtotal
    }

    # Two-decimal debug prints
    function Fmt2($n){ if ($null -ne $n) { ('{0:N2}' -f [decimal]$n) } else { '' } }
    Write-Output "   💳 Selection: $selectedMethod"
    if ($data.Totals.Subtotal)       { Write-Output "   💰 Subtotal: $(Fmt2 $data.Totals.Subtotal)" }
    if ($data.Totals.CashTotal)      { Write-Output "   💵 Cash Total Due: $(Fmt2 $data.Totals.CashTotal)" }
    if ($data.Totals.CreditTotal)    { Write-Output "   💳 Credit Total Due: $(Fmt2 $data.Totals.CreditTotal)" }
    if ($data.Totals.ConvenienceFee) { Write-Output "   🧾 Convenience Fee: $(Fmt2 $data.Totals.ConvenienceFee)" }
    Write-Output "   ✅ Using Total: $(Fmt2 $data.Totals.Total)"

    return $data
}

function Export-InvoiceData {
    <#
    .SYNOPSIS
        Exports parsed data in requested format
    #>
    param (
        [Parameter(Mandatory=$true)]
        [object]$Data,

        [Parameter(Mandatory=$true)]
        [string]$Format,

        [Parameter(Mandatory=$false)]
        [string]$OutputPath
    )

    switch ($Format) {
        "JSON" {
            $jsonPath = if ($OutputPath) { $OutputPath } else {
                ".\InvoiceData_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            }
            $Data | ConvertTo-Json -Depth 10 | Out-File $jsonPath -Encoding UTF8
            Write-Output "`nExported to JSON: $jsonPath"
            return $jsonPath
        }

        "Config" {
            $config = @{
                email_subject = "Freight Quote Request - $(if ($Data.PickupAddresses) { ($Data.PickupAddresses[0] -split ',')[0] } else { 'TBD' }) to Ashtabula, OH"
                auction_info = @{
                    auction_name = $Data.Vendor
                    pickup_address = if ($Data.PickupAddresses.Count -gt 0) { $Data.PickupAddresses[0] } else { "TBD" }
                    logistics_contact = @{
                        phone = if ($Data.ContactInfo.Phone.Count -gt 0) { $Data.ContactInfo.Phone[0] } else { "" }
                        email = if ($Data.ContactInfo.Email.Count -gt 0) { $Data.ContactInfo.Email[0] } else { "" }
                    }
                    pickup_datetime = if ($Data.PickupDates.Count -gt 0) { $Data.PickupDates[0] } else { "TBD" }
                    delivery_datetime = "TBD"
                    delivery_notice = "Driver must call at least one hour prior to delivery"
                    special_notes = $Data.SpecialNotes
                }
                delivery_address = "1218 Lake Avenue, Ashtabula, OH 44004"
                shipping_requirements = @{
                    total_pallets = "TBD"
                    truck_types = "TBD - Please recommend based on items"
                    labor_needed = "TBD"
                    weight_notes = "Total weight will NOT exceed standard truck capacity"
                }
            }

            $configPath = if ($OutputPath) { $OutputPath } else {
                ".\Config_$($Data.Vendor -replace '\s+','_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            }
            $config | ConvertTo-Json -Depth 10 | Out-File $configPath -Encoding UTF8
            Write-Output "`nExported logistics config: $configPath"
            return $configPath
        }

        default {
            # Display format
            Write-Output "`n╔════════════════════════════════════════════════════════╗"
            Write-Output "║           PARSED INVOICE DATA                          ║"
            Write-Output "╚════════════════════════════════════════════════════════╝"

            Write-Output "`nVENDOR INFORMATION"
            Write-Output "Vendor: $($Data.Vendor)"
            if ($Data.InvoiceNumber) {
                Write-Output "Invoice #: $($Data.InvoiceNumber)"
            }

            Write-Output "`nCONTACT INFORMATION"
            if ($Data.ContactInfo.Phone.Count -gt 0) {
                Write-Output "Phone(s):"
                foreach ($phone in $Data.ContactInfo.Phone) {
                    Write-Output "  • $phone"
                }
            }
            if ($Data.ContactInfo.Email.Count -gt 0) {
                Write-Output "Email(s):"
                foreach ($email in $Data.ContactInfo.Email) {
                    Write-Output "  • $email"
                }
            }

            if ($Data.PickupAddresses.Count -gt 0) {
                Write-Output "`nPICKUP ADDRESSES"
                foreach ($addr in $Data.PickupAddresses) {
                    Write-Output "• $addr"
                }
            }

            if ($Data.PickupDates.Count -gt 0) {
                Write-Output "`nPICKUP DATES"
                foreach ($date in $Data.PickupDates) {
                    Write-Output "• $date"
                }
            }

            if ($Data.Items.Count -gt 0) {
                Write-Output "`nITEMS ($($Data.Items.Count) lots)"
                $displayCount = [Math]::Min(10, $Data.Items.Count)
                for ($i = 0; $i -lt $displayCount; $i++) {
                    $item = $Data.Items[$i]
                    Write-Output "Lot $($item.LotNumber): $($item.Description)"
                }
                if ($Data.Items.Count -gt 10) {
                    Write-Output "... and $($Data.Items.Count - 10) more items"
                }
            }

            if ($Data.Totals.Total) {
                $totalFmt  = '{0:N2}' -f [decimal]$Data.Totals.Total
                Write-Output "`nTOTAL: `$$totalFmt"
                if ($Data.Totals.CashTotal -or $Data.Totals.CreditTotal) {
                    $cashFmt   = if ($Data.Totals.CashTotal)   { '{0:N2}' -f [decimal]$Data.Totals.CashTotal } else { '' }
                    $creditFmt = if ($Data.Totals.CreditTotal) { '{0:N2}' -f [decimal]$Data.Totals.CreditTotal } else { '' }
                    Write-Output "   (Cash Total: $cashFmt; Credit Total: $creditFmt)"
                }
            }

            if ($Data.SpecialNotes.Count -gt 0) {
                Write-Output "`nSPECIAL NOTES"
                foreach ($note in $Data.SpecialNotes) {
                    Write-Output "• $note"
                }
            }

            Write-Output ""
        }
    }
}
#endregion

#region Main Execution
Write-Output "`n╔════════════════════════════════════════════════════════╗"
Write-Output "║     GENERIC PDF INVOICE PARSER v1.3.2                 ║"
Write-Output "╚════════════════════════════════════════════════════════╝`n"

Initialize-InvoicePattern

if ($GUI) {
    Write-Output "GUI mode not yet implemented in this version"
    Write-Output "Use: .\Generic-PDF-Invoice-Parser.ps1 -PDFPath <path>"
}
elseif ($PDFPath) {
    if (!(Test-Path $PDFPath)) {
        Write-Error "File not found: $PDFPath"
        exit 1
    }

    # Check if it's a text file instead of PDF
    $extension = [System.IO.Path]::GetExtension($PDFPath).ToLower()
    
    if ($extension -eq '.txt') {
        Write-Output "`n📄 Reading text file..."
        $extraction = @{
            Text = Get-Content $PDFPath -Raw -Encoding UTF8
            Method = "TextFile"
            Quality = "High"
        }
    }
    else {
        # Extract text from PDF
        $extraction = Get-PDFTextContent -Path $PDFPath
    }

    if ($extraction -and $extraction.Text) {
        Write-Output "`n✅ Extraction successful:"
        Write-Output "   Method: $($extraction.Method)"
        Write-Output "   Quality: $($extraction.Quality)"
        Write-Output "   Length: $($extraction.Text.Length) characters"
        
        # Always save extracted text if in debug mode
        if ($DebugMode) {
            $debugFile = ".\DEBUG_ExtractedText_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $extraction.Text | Out-File -FilePath $debugFile -Encoding UTF8
            Write-Output "`n🔍 DEBUG: Saved extracted text to: $debugFile"
        }

        # Identify vendor and parse
        $vendorPattern = Find-InvoiceVendor -Text $extraction.Text
        $parsedData = Get-InvoiceData -Text $extraction.Text -VendorPattern $vendorPattern `
                                      -PaymentMethod $PaymentMethod -PromptPayment:$PromptPayment

        # Check if we got meaningful data
        $hasData = $parsedData.InvoiceNumber -or 
                   $parsedData.ContactInfo.Phone.Count -gt 0 -or 
                   $parsedData.ContactInfo.Email.Count -gt 0 -or
                   $parsedData.PickupAddresses.Count -gt 0 -or
                   $parsedData.Items.Count -gt 0

        if (-not $hasData -and -not $DebugMode) {
            Write-Output "`n⚠️  No data was extracted. Enabling debug mode..."
            $debugFile = ".\DEBUG_ExtractedText_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $extraction.Text | Out-File -FilePath $debugFile -Encoding UTF8
            Write-Output "📁 Saved extracted text to: $debugFile"
            Write-Output "`n💡 Please check the debug file to see what text was extracted from the PDF."
            Write-Output "   This will help determine if the PDF is readable or needs a different extraction method."
        }

        # Export in requested format
        Export-InvoiceData -Data $parsedData -Format $OutputFormat

        # Save debug data if requested
        if ($DebugMode) {
            $debugDataFile = ".\DEBUG_ParsedData_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $parsedData | ConvertTo-Json -Depth 10 | Out-File -FilePath $debugDataFile -Encoding UTF8
            Write-Output "`n🔍 DEBUG: Saved parsed data to: $debugDataFile"
        }
    }
    else {
        Write-Error "Failed to extract text from PDF"
        Write-Output "`n💡 QUICK FIX - Manual extraction:"
        Write-Output ""
        Write-Output "1. Open your PDF in Adobe Reader"
        Write-Output "2. Select All (Ctrl+A) and Copy (Ctrl+C)"
        Write-Output "3. Paste into Notepad and save as invoice.txt"
        Write-Output "4. Run: .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.txt"
        exit 1
    }
}
else {
    Write-Output "Usage: .\Generic-PDF-Invoice-Parser.ps1 -PDFPath <path> [-OutputFormat JSON|Config] [-DebugMode] [-PaymentMethod Cash|Credit] [-PromptPayment]"
    Write-Output ""
    Write-Output "Examples:"
    Write-Output "  .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.pdf"
    Write-Output "  .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.pdf -PaymentMethod Credit"
    Write-Output "  .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.pdf -PromptPayment"
}
#endregion
