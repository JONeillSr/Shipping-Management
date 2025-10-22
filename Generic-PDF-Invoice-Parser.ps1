<#
.SYNOPSIS
    Generic PDF Invoice Parser for multiple auction formats (Brolyn, Ritchie Bros, Stronghold Equipment, etc.).

.DESCRIPTION
    Extracts structured invoice data (vendor, buyer, items/lots, totals, fees, pickup info) from PDF invoices.
    Optimized for messy text extraction and inconsistent layouts. Includes a vendor-specific, stateful parser for
    Stronghold Equipment that pairs LOT → DESCRIPTION → HAMMER PRICE reliably even when lines wrap or prices
    appear detached (e.g., "$ 1,234.56" or "1,234.56$" on the next line).

.PARAMETER PDFPath
    Path to the invoice PDF.

.PARAMETER ReturnObject
    When set, returns a single PowerShell object with structured fields instead of writing console output.
    Intended for wrapper scripts (e.g., Invoice-ToPurchaseTrackingCSV.ps1) that generate CSVs.

.PARAMETER PaymentMethod
    Optional. 'Cash' or 'Credit'. Affects totals validation for layouts that show both "Cash Total Due" and
    "Credit Total Due". Defaults to 'Cash' unless the invoice clearly indicates otherwise.

.PARAMETER PromptPayment
    Optional switch. If present and the invoice indicates a prompt/early payment discount, the totals logic
    can factor that in when validating relationships.

.PARAMETER StrictTotals
    Optional switch. Enables strict cross-checks between Subtotal, Convenience Fee, and Grand/Cash/Credit totals.
    When enabled, the script throws if relationships are inconsistent rather than tolerating minor drift.

.OUTPUTS
    [pscustomobject]
    {
        Vendor, InvoiceNumber, InvoiceDate,
        Buyer: { Name, Address, Phone, Email },
        PickupAddresses: [ { Street, Address2, City, State, Zip, OneLine } ... ],
        PickupDates:     [ ... ],
        Items:           [ { LotNumber, Description, HammerPrice } ... ],
        Totals:          { Subtotal, ConvenienceFee, GrandTotal, CashTotal, CreditTotal }
    }

.EXAMPLE
    .\Generic-PDF-Invoice-Parser.ps1 -PDFPath .\SE Invoice-101625-5024.pdf

.EXAMPLE
    # Return a structured object for a wrapper script to convert to CSV
    $data = .\Generic-PDF-Invoice-Parser.ps1 -PDFPath .\SE Invoice-101625-5024.pdf -ReturnObject

.EXAMPLE
    # Strict totals validation assuming a credit-card checkout
    .\Generic-PDF-Invoice-Parser.ps1 -PDFPath .\invoice.pdf -PaymentMethod Credit -StrictTotals

.NOTES
    Version: v1.8.1
    Changes:
      - Stronghold: extraction rewritten as a small state machine to avoid lot/price misalignment and
        handle detached dollar signs, wide whitespace, and orphaned lot lines.
      - Keeps previous behavior for other vendors.
      - ReturnObject flow is unchanged and compatible with existing CSV wrappers.

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$PDFPath,

    [Parameter(Mandatory=$false)]
    [ValidateSet("JSON", "CSV", "Display", "Config")]
    [string]$OutputFormat = "Display",

    [Parameter(Mandatory=$false)]
    [switch]$ReturnObject,

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
    [switch]$PromptPayment,

    [Parameter(Mandatory=$false)]
    [switch]$StrictTotals
)

# Prevent ScriptAnalyzer PSReviewUnusedParameter warnings when code paths short-circuit:
$null = $PaymentMethod; $null = $PromptPayment; $null = $StrictTotals

# Explicit reference to satisfy analyzer (even if feature is a no-op today)
if ($SavePattern) { $null = $true }

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
        Uses pdftotext (xpdf-tools) if available, then iTextSharp as a fallback.
    #>
    param ([Parameter(Mandatory=$true)][string]$Path)

    if (-not $ReturnObject) {
        Write-Output "`n📄 Extracting text from PDF..."
        Write-Output "   File: $(Split-Path $Path -Leaf)"
    }

    $pdftotext = Get-Command pdftotext -ErrorAction SilentlyContinue
    if ($pdftotext) {
        try {
            $tempFile = [System.IO.Path]::GetTempFileName()
            & pdftotext -layout $Path $tempFile 2>&1 | Out-Null
            if (Test-Path $tempFile) {
                $extractedText = Get-Content $tempFile -Raw -Encoding UTF8
                Remove-Item $tempFile -Force
                if ($extractedText -and $extractedText.Length -gt 100) {
                    if (-not $ReturnObject) {
                        Write-Output "   ✅ Extracted $($extractedText.Length) characters using pdftotext"
                    }
                    return @{ Text = $extractedText; Method = "pdftotext"; Quality = "High" }
                }
            }
        } catch { Write-Verbose "pdftotext extraction failed: $_" }
    }

    try {
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
                if (-not $ReturnObject) {
                    Write-Output "   ✅ Extracted $($extractedText.Length) characters using iTextSharp"
                }
                return @{ Text = $extractedText; Method = "iTextSharp"; Quality = "High" }
            }
        }
    } catch { Write-Verbose "iTextSharp extraction failed: $_" }

    if (-not $ReturnObject) {
        Write-Warning "`n❌ Could not extract readable text from PDF"
        Write-Output "`n💡 This PDF requires a proper PDF parsing library."
        Write-Output "Option 1 - Install xpdf-tools (EASIEST): https://www.xpdfreader.com/download.html"
        Write-Output "Option 2 - Adobe Reader > Save As Other > Text, then pass the .txt file"
        Write-Output "Option 3 - Copy/paste to Notepad and save as .txt"
    }
    return $null
}
#endregion

#region Pattern Recognition Functions
function Import-InvoicePattern {
    <#
    .SYNOPSIS
        Loads learned patterns from previous parses
    #>
    if (Test-Path $script:PatternsFile) {
        $script:LearnedPatterns = Get-Content $script:PatternsFile -Raw | ConvertFrom-Json
        Write-Verbose "Loaded $($script:LearnedPatterns.Count) learned patterns"
    } else {
        $script:LearnedPatterns = @()

        $script:LearnedPatterns += @{
            Vendor = "Brolyn Auctions"
            Patterns = @{
                CompanyIdentifier = 'Brolyn|BROLYN'
                Phone = '\(574\)\s*891-3111'
                Email = 'logistics@brolynauctions\.com'
                Address = '290\s+West\s+750\s+North.*?Howe.*?IN|1139\s+Haines.*?Sturgis.*?MI'
                PickupDates = '(?:load times|pickup).*?(\w+\s+\d{1,2}/\d{1,2})\s+thru\s+(\w+\s+\d{1,2}/\d{1,2})'
                LotPattern = '^(\d{2,5})\s+\d{4}\s+(.+)'
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
                LotPattern = '^(\d{2,5})\s+\d{4}\s+(.+)'
            }
        }

        $script:LearnedPatterns += @{
            Vendor = "Stronghold Equipment"
            Patterns = @{
                CompanyIdentifier = 'STRONGHOLD EQUIPMENT|Stronghold Equipment|Rocky Recycling Ltd'
                Phone = '905-662-8200'
                Email = 'dan@strongholdequipment.com'
                Address = '349 Arvin Ave'
                PickupDates = 'Removal|Pickup'
            }
        }

    }

    if ($SavePattern) { $null = $script:LearnedPatterns.Count }
}

function Find-InvoiceVendor {
    <#
    .SYNOPSIS
        Identifies the auction vendor from invoice text with enhanced pattern matching
    #>
    param ([Parameter(Mandatory=$true)][string]$Text)

    if (-not $ReturnObject) { Write-Output "`n Identifying vendor..." }

    $cleanText = $Text -replace '\s+', ' '

    # Check for Stronghold Equipment
    if ($cleanText -match '(?i)stronghold|STRONGHOLD|Rocky Recycling') {
        if (-not $ReturnObject) { Write-Output "   Detected: Stronghold Equipment" }
        return $script:LearnedPatterns | Where-Object { $_.Vendor -eq "Stronghold Equipment" } | Select-Object -First 1
    }

    # Check for DC Auctions / Wavebid
    if ($cleanText -match '(?i)\bDC\s*Auctions\b|\bWavebid\b') {
        if (-not $ReturnObject) { Write-Output "   Detected: DC Auctions (Wavebid)" }
        return @{
            Vendor   = "DC Auctions"
            Patterns = @{
                CompanyIdentifier = 'DC\s*Auctions|Wavebid'
            }
        }
    }

    # Check for Brolyn
    if ($cleanText -match '(?i)brolyn|BROLYN') {
        if (-not $ReturnObject) { Write-Output "   Detected: Brolyn Auctions" }
        return $script:LearnedPatterns | Where-Object { $_.Vendor -eq "Brolyn Auctions" } | Select-Object -First 1
    }

    foreach ($pattern in $script:LearnedPatterns) {
        if ($cleanText -match $pattern.Patterns.CompanyIdentifier) {
            if (-not $ReturnObject) { Write-Output "   Detected: $($pattern.Vendor)" }
            return $pattern
        }
    }

    if (-not $ReturnObject) { Write-Output "   Unknown vendor - using generic patterns" }
    return @{
        Vendor = "Unknown"
        Patterns = @{
            CompanyIdentifier = ""
            Phone = '\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}'
            Email = '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
            Address = '\d+\s+[A-Za-z\s]+(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd|Drive|Dr|Lane|Ln)[,\s]+[A-Za-z\s]+[,\s]+[A-Z]{2}\s+\d{5}'
            PickupDates = '(\d{1,2}/\d{1,2}/\d{4})'
            LotPattern = '^(\d{2,5})\s+\d{4}\s+(.+)'
        }
    }
}

function Test-ValidPhoneNumber {
    <#
    .SYNOPSIS
        Validates if a phone number is legitimate (not a false positive)
    #>
    param ([Parameter(Mandatory=$true)][string]$Phone)

    $digits = $Phone -replace '[^\d]', ''
    if ($digits.Length -ne 10) { return $false }

    $areaCode = [int]$digits.Substring(0, 3)
    if ($areaCode -lt 200 -or $areaCode -gt 999) { return $false }

    $exchange = [int]$digits.Substring(3, 3)
    if ($exchange -lt 200 -or $exchange -gt 999) { return $false }

    return $true
}
#endregion

#region Address helpers (Address2 from parentheses)
function ConvertTo-AddressObject {
    <#
    .SYNOPSIS
        Builds a structured pickup address object and normalizes components.
    .DESCRIPTION
        If Street contains parenthetical content like "(Plant 208/209)", that
        text is moved to Address2. OneLine remains "Street, City ST ZIP".
    #>
    param(
        [Parameter(Mandatory=$true)][string]$Street,
        [Parameter(Mandatory=$true)][string]$City,
        [Parameter(Mandatory=$true)][string]$State,
        [Parameter(Mandatory=$true)][string]$Zip
    )
    $streetClean = ($Street -replace '\s+', ' ').Trim()
    $cityClean   = ($City   -replace '\s+', ' ').Trim()
    $stateClean  = ($State  -replace '\s+', '').Trim().ToUpper()
    $zipClean    = ($Zip    -replace '\s+', '').Trim()

    $addr2 = $null
    $mAddr = [regex]::Match($streetClean, '\(([^)]+)\)')
    if ($mAddr.Success) {
        $addr2 = $mAddr.Groups[1].Value.Trim()
        $streetClean = ($streetClean -replace '\([^)]+\)', '').Trim()
        $streetClean = ($streetClean -replace '\s{2,}', ' ')
    }

    $oneLine = "$streetClean, $cityClean $stateClean $zipClean"
    return [pscustomobject]@{
        Street   = $streetClean
        Address2 = $addr2
        City     = $cityClean
        State    = $stateClean
        Zip      = $zipClean
        OneLine  = $oneLine
    }
}

function ConvertFrom-FreeformAddress {
    <#
    .SYNOPSIS
        Parses a single-line freeform address into a structured object.
    .DESCRIPTION
        Tries comma style first, then space style. Extracts Address2 from street.
    #>
    param([Parameter(Mandatory=$true)][string]$AddressLine)

    $s = ($AddressLine -replace '\s+', ' ').Trim()

    # Comma style: "123 Main St (Plant 208/209), Howe IN 46746"
    $m = [regex]::Match($s, '^(?<street>.+?),\s*(?<city>[A-Za-z][A-Za-z\s.\''-]+)\s+(?<state>[A-Z]{2})\s+(?<zip>\d{5})$')
    if ($m.Success) {
        return ConvertTo-AddressObject -Street $m.Groups['street'].Value -City $m.Groups['city'].Value -State $m.Groups['state'].Value -Zip $m.Groups['zip'].Value
    }

    # Space style: "123 Main St (Plant 208/209) Howe IN 46746"
    $m = [regex]::Match($s, '^(?<street>.+?)\s+(?<city>[A-Za-z][A-Za-z\s.\''-]+)\s+(?<state>[A-Z]{2})\s+(?<zip>\d{5})$')
    if ($m.Success) {
        return ConvertTo-AddressObject -Street $m.Groups['street'].Value -City $m.Groups['city'].Value -State $m.Groups['state'].Value -Zip $m.Groups['zip'].Value
    }

    return $null
}

function Get-PickupLocationAddress {
    <#
    .SYNOPSIS
        Extracts pickup addresses from "Location:" lines within the raw text.
    .DESCRIPTION
        Captures street (incl. optional parenthetical like "(Plant 208/209)"),
        city, state, ZIP. Returns de-duplicated [pscustomobject[]].
        NOTE: Uses a Hashtable for de-dupe (avoids generic HashSet[string]).
    #>
    param([Parameter(Mandatory=$true)][string]$RawText)

    $results = New-Object System.Collections.ArrayList
    $seen    = @{}   # key -> $true

    # Pattern A: comma between street and city
    $patternA = "Location:\s*(?<street>\d+[^\r\n,]*?)\s*,\s*(?<city>[A-Za-z][A-Za-z\s.\'-]+)\s+(?<state>[A-Z]{2})\s+(?<zip>\d{5})"
    foreach ($match in [System.Text.RegularExpressions.Regex]::Matches($RawText, $patternA, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
        $obj = ConvertTo-AddressObject -Street $match.Groups['street'].Value -City $match.Groups['city'].Value -State $match.Groups['state'].Value -Zip $match.Groups['zip'].Value
        $addr2Key = if ($null -ne $obj.Address2 -and $obj.Address2 -ne '') { $obj.Address2 } else { '' }
        $key = [string]::Format("{0}|{1}", $obj.OneLine, $addr2Key)
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $null = $results.Add($obj)
        }
    }

    # Pattern B: no comma between street and city (fallback)
    $patternB = "Location:\s*(?<street>\d+[^\r\n,]*?)\s+(?<city>[A-Za-z][A-Za-z\s.\'-]+)\s+(?<state>[A-Z]{2})\s+(?<zip>\d{5})"
    foreach ($match in [System.Text.RegularExpressions.Regex]::Matches($RawText, $patternB, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
        $obj = ConvertTo-AddressObject -Street $match.Groups['street'].Value -City $match.Groups['city'].Value -State $match.Groups['state'].Value -Zip $match.Groups['zip'].Value
        $addr2Key = if ($null -ne $obj.Address2 -and $obj.Address2 -ne '') { $obj.Address2 } else { '' }
        $key = [string]::Format("{0}|{1}", $obj.OneLine, $addr2Key)
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $null = $results.Add($obj)
        }
    }

    return $results.ToArray()
}

#endregion

function Get-InvoiceData {
    <#
    .SYNOPSIS
        Extracts structured data from invoice text with enhanced pattern matching
    .DESCRIPTION
        Parses vendor, contacts, pickup info (structured), items, and totals.
        Supports vendor-specific item extraction including Hammer Price for Stronghold.
        Totals logic:
          - Windowed character search after labels (prevents column drift).
          - Relationship rules: Cash = Subtotal; Credit = Subtotal + Convenience Fee.
          - -PromptPayment asks user Cash/Credit when possible.
          - -StrictTotals throws on ambiguity or conflicts.
    #>
    param (
        [Parameter(Mandatory=$true)][string]$Text,
        [Parameter(Mandatory=$true)][object]$VendorPattern,
        [Parameter(Mandatory=$true)][ValidateSet("Cash","Credit")][string]$PaymentMethod,
        [Parameter(Mandatory=$true)][bool]$PromptPayment,
        [Parameter(Mandatory=$true)][bool]$StrictTotals
    )

    if (-not $ReturnObject) { Write-Output "`n📊 Extracting invoice data..." }

    $normalizedText = $Text -replace '[\r\n]+', ' '
    $normalizedText = $normalizedText -replace '\s+', ' '

    function Fmt2($n){ if ($null -ne $n) { ('{0:N2}' -f [decimal]$n) } else { '' } }

    $data = [ordered]@{
        Vendor = $VendorPattern.Vendor
        InvoiceNumber = $null
        InvoiceDate = $null
        ContactInfo = @{ Phone = @(); Email = @() }
        PickupAddresses = @()   # structured objects
        PickupDates = @()
        Items = @()
        Totals = @{
            Subtotal = $null; Tax = $null; Premium = $null; Total = $null
            CashTotal = $null; CreditTotal = $null; ConvenienceFee = $null
        }
        SpecialNotes = @()
    }
    # Ensure Items holds PSCustomObjects and avoid duplicates across parsing passes
    if (-not ($data.Items -is [System.Collections.IList])) { $data.Items = @() }
    $seenKeys = New-Object 'System.Collections.Generic.HashSet[string]'

    function Commit-Item([string]$lot, [string]$desc, [string]$price) {
        if (-not $lot)   { return }
        if (-not $price) { return }
        $k = "$lot|$price"
        if ($seenKeys.Contains($k)) { return }
        $null = $seenKeys.Add($k)
        $data.Items += [pscustomobject]@{
            LotNumber   = $lot
            Description = ($desc | ForEach-Object { $_ }) -as [string]
            HammerPrice = $price
        }
    }


    # Invoice #
    if ($normalizedText -match '(?:Invoice\s*#?\s*:?\s*)?(\d{4}-\d{6}-\d+)') {
        $data.InvoiceNumber = $Matches[1]
        if (-not $ReturnObject) { Write-Output "   📋 Invoice #: $($data.InvoiceNumber)" }
    }
    elseif ($normalizedText -match 'Invoice\s*#?\s*:?\s*([A-Z0-9-]+)') {
        $data.InvoiceNumber = $Matches[1]
        if (-not $ReturnObject) { Write-Output "   📋 Invoice #: $($data.InvoiceNumber)" }
    }
    elseif ($normalizedText -match 'BUYER/INVOICE\s+(\d+)') {
        $data.InvoiceNumber = $Matches[1]
        if (-not $ReturnObject) { Write-Output "   📋 Invoice #: $($data.InvoiceNumber)" }
    }

    # Date
    if ($normalizedText -match 'Date:\s*(\d{1,2}/\d{1,2}/\d{4})') {
        $data.InvoiceDate = $Matches[1]
        if (-not $ReturnObject) { Write-Output "   📅 Date: $($data.InvoiceDate)" }
    }
    elseif ($normalizedText -match 'Invoice\s+Date\s*:?\s*(\d{1,2}[-/]\w+[-/]\d{4}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})') {
        $data.InvoiceDate = $Matches[1]
        if (-not $ReturnObject) { Write-Output "   📅 Date: $($data.InvoiceDate)" }
    }
    elseif ($normalizedText -match '\d{2}-\w{3}-\d{4}\s+\d{2}:\d{2}') {
        $data.InvoiceDate = $Matches[0]
        if (-not $ReturnObject) { Write-Output "   📅 Date: $($data.InvoiceDate)" }
    }
    elseif ($normalizedText -match 'DATE:\s*(\d{1,2}/\d{1,2}/\d{4})') {
        $data.InvoiceDate = $Matches[1]
        if (-not $ReturnObject) { Write-Output "   📅 Date: $($data.InvoiceDate)" }
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
    if (-not $ReturnObject -and $data.ContactInfo.Phone.Count -gt 0) {
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
    if (-not $ReturnObject -and $data.ContactInfo.Email.Count -gt 0) {
        Write-Output "   📧 Found $($data.ContactInfo.Email.Count) email(s)"
    }

    # Generic address hits (convert to structured objects)
    $addressPatterns = @(
        '(\d+\s+(?:West|East|North|South)\s+\d+\s+(?:West|East|North|South)[^\n,]+?(?:Howe|Sturgis)[^\n,]+?(?:IN|MI)\s+\d{5})',
        '(\d+\s+[A-Za-z]+\s+(?:Blvd|Boulevard|Street|St|Avenue|Ave|Road|Rd|Drive|Dr)[^\n,]+?(?:Howe|Sturgis)[^\n,]+?(?:IN|MI)\s+\d{5})',
        '(\d+\s+[A-Za-z]+\s+(?:Blvd|Boulevard|Street|St|Avenue|Ave|Road|Rd|Drive|Dr)[^\n,]+?(?:Stoney\s+Creek)[^\n,]+?(?:ON)\s+[A-Z0-9]{3}\s+[A-Z0-9]{3})'
    )
    foreach ($pattern in $addressPatterns) {
        $addressMatches = [regex]::Matches($Text, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($m in $addressMatches) {
            $line = ($m.Groups[1].Value.Trim() -replace '\s+', ' ')
            $obj = ConvertFrom-FreeformAddress -AddressLine $line
            if ($null -ne $obj) {
                if (-not ($data.PickupAddresses | Where-Object { $_.OneLine -eq $obj.OneLine -and $_.Address2 -eq $obj.Address2 })) {
                    $data.PickupAddresses += $obj
                }
            }
        }
    }

    # Addresses from "Location:" lines
    $locationObjs = Get-PickupLocationAddress -RawText $Text
    foreach ($obj in $locationObjs) {
        if (-not ($data.PickupAddresses | Where-Object { $_.OneLine -eq $obj.OneLine -and $_.Address2 -eq $obj.Address2 })) {
            $data.PickupAddresses += $obj
        }
    }
    if (-not $ReturnObject -and $data.PickupAddresses.Count -gt 0) {
        Write-Output "   📍 Found $($data.PickupAddresses.Count) pickup address(es)"
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
    if (-not $ReturnObject -and $data.PickupDates.Count -gt 0) {
        Write-Output "   📅 Found $($data.PickupDates.Count) pickup date range(s)"
    }

    # --- DC Auctions / Wavebid parser ---
if ($VendorPattern.Vendor -eq "DC Auctions") {
    $lines = ($Text -split "`n") | ForEach-Object { $_.TrimEnd() }
    $n = $lines.Count

    # Table header + row patterns
    $headerPattern = '^(?i)Lot\s+Paddle\s+Description\s+Qty\s+Bid\s+Sale\s*Price'
    $lotFull       = '^(?<lot>\d+[A-Z]?)\b(?:\s+.*)?$'      # new lot begins when a line starts with 113, 434, 449A, etc.
    $priceLine     = '(\$[ \t]*\d{1,3}(?:,\d{3})*\.\d{2})'  # capture $x,xxx.xx
    $loadingFeeKW  = '(?i)\bLoading\s*Fee\b'

    function Get-FirstSalePriceFromLine([string]$s) {
        # prefer lines that look like right-column totals: multiple $ amounts; take the FIRST
        $mLocal = [regex]::Matches($s, $priceLine)
        if ($matches.Count -ge 1) {
            # first $-amount is the Sale Price for Wavebid invoice rows
            return ($matches[0].Value -replace '[^\d\.]', '') -replace ',', ''
        }
        return $null
    }

    $inTable     = $false
    $currentLot  = $null
    $currentDesc = ""

    for ($i=0; $i -lt $n; $i++) {
        $line = $lines[$i].Trim()

        if (-not $inTable) {
            if ($line -match $headerPattern) { $inTable = $true; continue }
            continue
        }

        # End of items area markers
        if ($line -match '^(?i)Total\s+Lots:\s*\d+' -or $line -match '^(?i)Totals:') {
            if ($currentLot) {
                # Try to flush pending lot using any trailing price we saw in description or next lines
                $flushed = $false
                for ($j=$i; $j -lt [Math]::Min($i+4,$n); $j++) {
                    $pl = $lines[$j]
                    if ($pl -match $loadingFeeKW) { continue }
                    $p = Get-FirstSalePriceFromLine $pl
                    if ($p) {
                        Commit-Item $currentLot ($currentDesc.Trim()) $p
                        $flushed = $true; break
                    }
                }
                if (-not $flushed) {
                    $p = Get-FirstSalePriceFromLine $currentDesc
                    if ($p) { Commit-Item $currentLot ($currentDesc.Trim()) $p }
                }
                $currentLot=$null; $currentDesc=""
            }
            break
        }

        # New lot row begins
        if ($line -match $lotFull) {
            # Flush previous pending lot first
            if ($currentLot) {
                # look ahead for a price line that is not a loading-fee row
                $added = $false
                for ($k=$i; $k -lt [Math]::Min($i+4,$n); $k++) {
                    $look = $lines[$k]
                    if ($look -match $loadingFeeKW) { continue }
                    $p = Get-FirstSalePriceFromLine $look
                    if ($p) {
                        Commit-Item $currentLot ($currentDesc.Trim()) $p
                        $added = $true; break
                    }
                }
                if (-not $added) {
                    # fallback: try inside the description
                    $p = Get-FirstSalePriceFromLine $currentDesc
                    if ($p) { Commit-Item $currentLot ($currentDesc.Trim()) $p }
                }
            }

            $currentLot  = $Matches['lot']
            $currentDesc = ""   # description accumulates from subsequent lines
            continue
        }

        # While inside a lot, accumulate description and try to capture Sale Price
        if ($currentLot) {
            # Skip pure "Loading Fee" rows from affecting price or description
            if ($line -match $loadingFeeKW) { continue }

            # If this line has multiple $ amounts (Sale Price / Premium / Tax / Total), take the first as Sale Price
            $p = Get-FirstSalePriceFromLine $line
            if ($p) {
                # Commit the lot immediately when we see the first $ amount line (Sale Price)
                $descOnly = $currentDesc.Trim()
                if (-not $descOnly) { $descOnly = $line -replace $priceLine, '' }
                Commit-Item $currentLot ($descOnly.Trim()) $p
                $currentLot  = $null
                $currentDesc = ""
            }
            else {
                # otherwise accumulate human-readable description lines
                if ($line) {
                    if ($currentDesc) { $currentDesc += " " + $line.Trim() } else { $currentDesc = $line.Trim() }
                }
            }
        }
    }

    # Safety flush in case file ended without totals block
    if ($currentLot) {
        $p = Get-FirstSalePriceFromLine $currentDesc
        if ($p) { Commit-Item $currentLot ($currentDesc.Trim()) $p }
        $currentLot=$null; $currentDesc=""
    }
}
# --- end DC Auctions / Wavebid parser ---

if ($VendorPattern.Vendor -eq "Stronghold Equipment") {
        # Stronghold-specific extraction (state machine; robust to messy OCR and line breaks)
        $lines = ($Text -split "`n") | ForEach-Object { $_.TrimEnd() }

        # A price can be "$ 1,234.56", "1,234.56$", or "1,234.56"
        $pricePattern = '(?<!\d)(?:\$\s*)?\d{1,3}(?:,\d{3})*\.\d{2}(?:\s*\$)?(?!\d)'
        # Lot number can be like "257A"
        $lotPattern   = '^(?<lot>\d+[A-Z]?)\b\s*(?<rest>.*)$'

        $inItems = $false
        $currentLot = $null
        $currentDesc = ""

        function Add-ItemIfHasPrice([string]$lot, [string]$text) {
            if (-not $lot) { return $false }
            $m = [regex]::Match($text, $pricePattern)
            if ($m.Success) {
                $price = ($m.Value -replace '[^\d\.]', '')  # strip $/commas/spaces
                $desc  = ($text -replace $pricePattern, '').Trim()
                if (-not $desc -or $desc.Length -lt 3) { $desc = $text.Trim() } # fallback
                $data.Items += @{
                    LotNumber   = $lot
                    Description = $desc
                    HammerPrice = $price
                }
                return $true
            }
            return $false
        }

        foreach ($rawLine in $lines) {
            $line = $rawLine.Trim()

            # Start item parsing when we see a Lot header *or* the first lot row
            if ($line -match '^(?i)Lot#') { $inItems = $true; continue }
            if (-not $inItems -and $line -match $lotPattern) { $inItems = $true }

            if (-not $inItems) { continue }
                if ($line -match '^\*\*\*|^Subtotal\b|^Notes:') {
                if ($currentLot) { $null = Add-ItemIfHasPrice $currentLot $currentDesc | Out-Null }
                break
            }

            if ($line -match $lotPattern) {
                if ($currentLot) { $null = Add-ItemIfHasPrice $currentLot $currentDesc | Out-Null }
                $currentLot  = $Matches['lot']
                $currentDesc = $Matches['rest'].Trim()
                if (Add-ItemIfHasPrice $currentLot $currentDesc) { $currentLot = $null; $currentDesc = "" }
                continue
            }

            if ($currentLot) {
                $candidate = if ($currentDesc) { "$currentDesc $line" } else { $line }
                if (Add-ItemIfHasPrice $currentLot $candidate) {
                    $currentLot  = $null
                    $currentDesc = ""
                } elseif ($line -notmatch '^\d+[A-Z]?$') {
                    $currentDesc = $candidate.Trim()
                }
            }
        }

        # Drop incomplete pending lot if no price was ever found
        $currentLot = $null
        $currentDesc = ""
} else {
        # Original Brolyn/Ritchie logic
        $lotLines = $Text -split "`n"
        $currentLot = $null
        foreach ($lineRaw in $lotLines) {
            $line = $lineRaw.Trim()
            $matchPattern = '^\d{2,5}\s+\d{4}\s+.+'
            if ($line -match $matchPattern) {
                if ($line -match '^(\d{2,5})\s+\d{4}\s+(.+)') {
                    $lotNum = $Matches[1]
                    $desc = $Matches[2].Trim()
                    if ($desc.Length -ge 15 -and $desc.Length -le 500) {
                        $desc = $desc -replace '\s+', ' '
                        $data.Items += @{ LotNumber = $lotNum; Description = $desc }
                        $currentLot = @{ LotNumber = $lotNum; Description = $desc }
                    }
                }
            }
            elseif ($null -ne $currentLot -and $line.Length -gt 10 -and $line.Length -lt 200) {
                $idx = $data.Items.Count - 1
                if ($idx -ge 0) {
                    $data.Items[$idx].Description += " " + $line.Trim()
                }
            }
        }
    }

    if (-not $ReturnObject -and $data.Items.Count -gt 0) {
        Write-Output "   📦 Found $($data.Items.Count) lot item(s)"
    }

    # ===== Totals =====
    $norm = ($Text -replace '\s+', ' ').Trim()
    function ConvertTo-Amount([string]$s) { if ($s) { [decimal]($s -replace ',','') } else { $null } }
    function NearEqual([decimal]$a, [decimal]$b, [decimal]$eps = 0.01) { return [Math]::Abs($a - $b) -le $eps }

    function GetAmountAfterLabel {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)][string]$Text,
            [Parameter(Mandatory=$true)][string]$Label,
            [Parameter(Mandatory=$false)][int]$Window = 100
        )
        $m = [regex]::Match($Text, [regex]::Escape($Label), 'IgnoreCase')
        if (-not $m.Success) { return $null }
        $start = [Math]::Min($Text.Length, $m.Index + $m.Length)
        $len   = [Math]::Min($Window, [Math]::Max(0, $Text.Length - $start))
        if ($len -le 0) { return $null }
        $slice = $Text.Substring($start, $len)
        $m2 = [regex]::Match($slice, '\$\s*([0-9,]+\.[0-9]{2})')
        if ($m2.Success) { return (ConvertTo-Amount -s $m2.Groups[1].Value) }
        return $null
    }

    $capturedSubtotal       = GetAmountAfterLabel -Text $norm -Label 'SubTotal:'         -Window 80
    $capturedCashTotal      = GetAmountAfterLabel -Text $norm -Label 'Cash Total Due:'   -Window 80
    $capturedConvenienceFee = GetAmountAfterLabel -Text $norm -Label 'Convenience Fee'   -Window 80
    $capturedCreditTotal    = GetAmountAfterLabel -Text $norm -Label 'Credit Total Due:' -Window 80
    $capturedGrandTotal     = GetAmountAfterLabel -Text $norm -Label 'Grand Total:'      -Window 80

    # Also try "Total in US Dollars" for Stronghold
    if ($null -eq $capturedGrandTotal) {
        $capturedGrandTotal = GetAmountAfterLabel -Text $norm -Label 'Total in US Dollars' -Window 80
    }

    # Try "Buyer's Premium" for Stronghold
    if ($null -eq $capturedConvenienceFee) {
        $capturedConvenienceFee = GetAmountAfterLabel -Text $norm -Label "Buyer's Premium" -Window 80
    }

    if ($null -ne $capturedSubtotal)       { $data.Totals.Subtotal       = $capturedSubtotal }
    if ($null -ne $capturedConvenienceFee) { $data.Totals.ConvenienceFee = $capturedConvenienceFee; $data.Totals.Premium = $capturedConvenienceFee }
    if ($null -ne $capturedCashTotal)      { $data.Totals.CashTotal      = $capturedCashTotal }
    if ($null -ne $capturedCreditTotal)    { $data.Totals.CreditTotal    = $capturedCreditTotal }

    if ($null -ne $data.Totals.Subtotal) {
        $sub = [decimal]$data.Totals.Subtotal
        $fee = if ($null -ne $data.Totals.ConvenienceFee) { [decimal]$data.Totals.ConvenienceFee } else { $null }

        if ($null -eq $data.Totals.CashTotal -or ($data.Totals.CashTotal -lt $sub)) { $data.Totals.CashTotal = $sub }
        if ($null -ne $fee -and $data.Totals.CashTotal -eq $fee) { $data.Totals.CashTotal = $sub }

        if ($null -ne $fee) {
            $data.Totals.CreditTotal = $sub + $fee
        }
        elseif ($null -eq $data.Totals.CreditTotal -and $null -ne $capturedGrandTotal) {
            $data.Totals.CreditTotal = $null
            if ($null -eq $data.Totals.CashTotal) { $data.Totals.CashTotal = $capturedGrandTotal }
        }
    }

    $selectedMethod = $PaymentMethod
    if ($PromptPayment -and $null -ne $data.Totals.CashTotal -and ($null -ne $data.Totals.CreditTotal -or $null -ne $data.Totals.ConvenienceFee)) {
        $answer = Read-Host "Payment method for totals (Cash/Credit) [$selectedMethod]"
        if     ($answer -match '^(?i)credit$') { $selectedMethod = 'Credit' }
        elseif ($answer -match '^(?i)cash$')   { $selectedMethod = 'Cash' }
    }

    if ($StrictTotals) {
        if ($null -eq $data.Totals.Subtotal) { throw "StrictTotals: Missing Subtotal; cannot determine totals confidently." }
        $sub = [decimal]$data.Totals.Subtotal
        $fee = if ($null -ne $data.Totals.ConvenienceFee) { [decimal]$data.Totals.ConvenienceFee } else { $null }

        switch ($selectedMethod) {
            'Credit' {
                if ($null -ne $fee) {
                    $calcCredit = $sub + $fee
                    if ($null -ne $capturedCreditTotal -and -not (NearEqual -a $capturedCreditTotal -b $calcCredit)) {
                        throw ("StrictTotals: Captured Credit Total ({0}) disagrees with Subtotal+Fee ({1})." -f (Fmt2 $capturedCreditTotal),(Fmt2 $calcCredit))
                    }
                }
                elseif ($null -eq $capturedCreditTotal) {
                    throw "StrictTotals: Credit total requires either a captured 'Credit Total Due' or both Subtotal and Convenience Fee."
                }
            }
            default {
                if ($null -ne $capturedCashTotal) {
                    if (-not (NearEqual -a $capturedCashTotal -b $sub)) {
                        if ($null -ne $fee -and (NearEqual -a $capturedCashTotal -b $fee)) {
                            throw ("StrictTotals: Captured Cash Total equals the Convenience Fee ({0}); ambiguous layout." -f (Fmt2 $capturedCashTotal))
                        } else {
                            throw ("StrictTotals: Captured Cash Total ({0}) disagrees with Subtotal ({1})." -f (Fmt2 $capturedCashTotal),(Fmt2 $sub))
                        }
                    }
                }
            }
        }
    }

    switch ($selectedMethod) {
        'Credit' {
            if     ($null -ne $data.Totals.CreditTotal)                      { $data.Totals.Total = $data.Totals.CreditTotal }
            elseif ($null -ne $data.Totals.Subtotal -and $null -ne $fee)     { $data.Totals.Total = [decimal]$data.Totals.Subtotal + [decimal]$fee }
            else                                                             { $data.Totals.Total = $data.Totals.CashTotal }
        }
        default {
            $data.Totals.Total = if ($null -ne $capturedGrandTotal) { $capturedGrandTotal } else { $data.Totals.CashTotal }
        }
    }

    if (-not $StrictTotals -and $null -ne $data.Totals.Total -and $null -ne $data.Totals.Subtotal -and $data.Totals.Total -lt $data.Totals.Subtotal) {
        if (-not $ReturnObject) { Write-Warning "Total ($($data.Totals.Total)) < Subtotal ($($data.Totals.Subtotal)); correcting to Subtotal for Cash." }
        $data.Totals.Total = $data.Totals.Subtotal
        $data.Totals.CashTotal = $data.Totals.Subtotal
    }

    if (-not $ReturnObject) {
        Write-Output "   💳 Selection: $selectedMethod"
        if ($null -ne $data.Totals.Subtotal)       { Write-Output "   💰 Subtotal: $(Fmt2 $data.Totals.Subtotal)" }
        if ($null -ne $data.Totals.CashTotal)      { Write-Output "   💵 Cash Total Due: $(Fmt2 $data.Totals.CashTotal)" }
        if ($null -ne $data.Totals.CreditTotal)    { Write-Output "   💳 Credit Total Due: $(Fmt2 $data.Totals.CreditTotal)" }
        if ($null -ne $data.Totals.ConvenienceFee) { Write-Output "   🧾 Convenience Fee: $(Fmt2 $data.Totals.ConvenienceFee)" }
        Write-Output "   ✅ Using Total: $(Fmt2 $data.Totals.Total)"
    }

    return [pscustomobject]$data

}
function Export-InvoiceData {
    <#
    .SYNOPSIS
        Exports or prints parsed data in requested format (or returns path).
    #>
    param (
        [Parameter(Mandatory=$true)][object]$Data,
        [Parameter(Mandatory=$true)][string]$Format,
        [Parameter(Mandatory=$false)][string]$OutputPath
    )

    if ($ReturnObject) { return $Data }

    switch ($Format) {
        "JSON" {
            $jsonPath = if ($OutputPath) { $OutputPath } else { ".\InvoiceData_$(Get-Date -Format 'yyyyMMdd_HHmmss').json" }
            $Data | ConvertTo-Json -Depth 10 | Out-File $jsonPath -Encoding UTF8
            Write-Output "`nExported to JSON: $jsonPath"
            return $jsonPath
        }
        "Config" {
            # Multiline pickup address string
            $pickupMultiline = if ($Data.PickupAddresses.Count -gt 0) {
                $a = $Data.PickupAddresses[0]
                if ($a.Address2) {
                    "$($a.Street)`n$($a.Address2)`n$($a.City) $($a.State) $($a.Zip)"
                } else {
                    "$($a.Street)`n$($a.City) $($a.State) $($a.Zip)"
                }
            } else { "TBD" }

            $config = @{
                email_subject = "Freight Quote Request - $(
                    if ($Data.PickupAddresses.Count -gt 0) { $Data.PickupAddresses[0].Street } else { 'TBD' }
                ) to Ashtabula, OH"
                auction_info = @{
                    auction_name   = $Data.Vendor
                    pickup_address = $pickupMultiline
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
                    truck_types   = "TBD - Please recommend based on items"
                    labor_needed  = "TBD"
                    weight_notes  = "Total weight will NOT exceed standard truck capacity"
                }
            }
            $configPath = if ($OutputPath) { $OutputPath } else { ".\Config_$($Data.Vendor -replace '\s+','_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').json" }
            $config | ConvertTo-Json -Depth 10 | Out-File $configPath -Encoding UTF8
            Write-Output "`nExported logistics config: $configPath"
            return $configPath
        }
        default {
            Write-Output "`n╔════════════════════════════════════════════════════════╗"
            Write-Output "║           PARSED INVOICE DATA                          ║"
            Write-Output "╚════════════════════════════════════════════════════════╝"

            Write-Output "`nVENDOR INFORMATION"
            Write-Output "Vendor: $($Data.Vendor)"
            if ($Data.InvoiceNumber) { Write-Output "Invoice #: $($Data.InvoiceNumber)" }

            Write-Output "`nCONTACT INFORMATION"
            if ($Data.ContactInfo.Phone.Count -gt 0) {
                Write-Output "Phone(s):"
                foreach ($p in $Data.ContactInfo.Phone) { Write-Output "  • $p" }
            }
            if ($Data.ContactInfo.Email.Count -gt 0) {
                Write-Output "Email(s):"
                foreach ($e in $Data.ContactInfo.Email) { Write-Output "  • $e" }
            }

            if ($Data.PickupAddresses.Count -gt 0) {
                Write-Output "`nPICKUP ADDRESSES"
                foreach ($a in $Data.PickupAddresses) {
                    if ($a.Address2) {
                        Write-Output "• $($a.Street)"
                        Write-Output "  $($a.Address2)"
                        Write-Output "  $($a.City) $($a.State) $($a.Zip)"
                    } else {
                        Write-Output "• $($a.Street)"
                        Write-Output "  $($a.City) $($a.State) $($a.Zip)"
                    }
                }
            }

            if ($Data.PickupDates.Count -gt 0) {
                Write-Output "`nPICKUP DATES"
                foreach ($d in $Data.PickupDates) { Write-Output "• $d" }
            }

            if ($Data.Items.Count -gt 0) {
                Write-Output "`nITEMS ($($Data.Items.Count) lots)"
                $displayCount = [Math]::Min(10, $Data.Items.Count)
                for ($i = 0; $i -lt $displayCount; $i++) {
                    $item = $Data.Items[$i]
                    $priceInfo = if ($item.HammerPrice) { " - `$$($item.HammerPrice)" } else { "" }
                    Write-Output "Lot $($item.LotNumber): $($item.Description)$priceInfo"
                }
                if ($Data.Items.Count -gt 10) { Write-Output "... and $($Data.Items.Count - 10) more items" }
            }

            if ($Data.Totals.Total) {
                $totalFmt = '{0:N2}' -f [decimal]$Data.Totals.Total
                Write-Output "`nTOTAL: `$$totalFmt"
                if ($Data.Totals.CashTotal -or $Data.Totals.CreditTotal) {
                    $cashFmt   = if ($Data.Totals.CashTotal)   { '{0:N2}' -f [decimal]$Data.Totals.CashTotal } else { '' }
                    $creditFmt = if ($Data.Totals.CreditTotal) { '{0:N2}' -f [decimal]$Data.Totals.CreditTotal } else { '' }
                    Write-Output "   (Cash Total: $cashFmt; Credit Total: $creditFmt)"
                }
            }

            if ($Data.SpecialNotes.Count -gt 0) {
                Write-Output "`nSPECIAL NOTES"
                foreach ($note in $Data.SpecialNotes) { Write-Output "• $note" }
            }

            Write-Output ""
        }
    }
}
#endregion

#region Main Execution
if (-not $ReturnObject) {
    Write-Output "`n╔════════════════════════════════════════════════════════╗"
    Write-Output "║     GENERIC PDF INVOICE PARSER v1.7.0                 ║"
    Write-Output "╚════════════════════════════════════════════════════════╝`n"
}

Import-InvoicePattern

if ($GUI) {
    if (-not $ReturnObject) {
        Write-Output "GUI mode not yet implemented in this version"
        Write-Output "Use: .\Generic-PDF-Invoice-Parser.ps1 -PDFPath <path>"
    }
}
elseif ($PDFPath) {
    if (-not (Test-Path $PDFPath)) {
        Write-Error "File not found: $PDFPath"
        if ($ReturnObject) { return $null } else { exit 1 }
    }

    $extension = [System.IO.Path]::GetExtension($PDFPath).ToLower()
    if ($extension -eq '.txt') {
        if (-not $ReturnObject) { Write-Output "`n📄 Reading text file..." }
        $extraction = @{ Text = Get-Content $PDFPath -Raw -Encoding UTF8; Method = "TextFile"; Quality = "High" }
    } else {
        $extraction = Get-PDFTextContent -Path $PDFPath
    }

    if ($extraction -and $extraction.Text) {
        if (-not $ReturnObject) {
            Write-Output "`n✅ Extraction successful:"
            Write-Output "   Method: $($extraction.Method)"
            Write-Output "   Quality: $($extraction.Quality)"
            Write-Output "   Length: $($extraction.Text.Length) characters"
        }

        if ($DebugMode) {
            $debugFile = ".\DEBUG_ExtractedText_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $extraction.Text | Out-File -FilePath $debugFile -Encoding UTF8
            if (-not $ReturnObject) { Write-Output "`n🔍 DEBUG: Saved extracted text to: $debugFile" }
        }

        $vendorPattern = Find-InvoiceVendor -Text $extraction.Text
        $parsedData = Get-InvoiceData -Text $extraction.Text -VendorPattern $vendorPattern `
                                      -PaymentMethod $PaymentMethod -PromptPayment:$PromptPayment `
                                      -StrictTotals:$StrictTotals

        $hasData = $parsedData.InvoiceNumber -or
                   $parsedData.ContactInfo.Phone.Count -gt 0 -or
                   $parsedData.ContactInfo.Email.Count -gt 0 -or
                   $parsedData.PickupAddresses.Count -gt 0 -or
                   $parsedData.Items.Count -gt 0

        if (-not $hasData -and -not $DebugMode -and -not $ReturnObject) {
            Write-Output "`n⚠️  No data was extracted. Enabling debug mode..."
            $debugFile = ".\DEBUG_ExtractedText_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $extraction.Text | Out-File -FilePath $debugFile -Encoding UTF8
            Write-Output "📁 Saved extracted text to: $debugFile"
            Write-Output "`n💡 Please check the debug file to see what text was extracted from the PDF."
        }

        if ($ReturnObject) { return $parsedData }

        Export-InvoiceData -Data $parsedData -Format $OutputFormat

        if ($DebugMode) {
            $debugDataFile = ".\DEBUG_ParsedData_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $parsedData | ConvertTo-Json -Depth 10 | Out-File -FilePath $debugDataFile -Encoding UTF8
            if (-not $ReturnObject) { Write-Output "`n🔍 DEBUG: Saved parsed data to: $debugDataFile" }
        }
    }
    else {
        Write-Error "Failed to extract text from PDF"
        if (-not $ReturnObject) {
            Write-Output "`n💡 QUICK FIX - Manual extraction:"
            Write-Output "1. Open your PDF in Adobe Reader"
            Write-Output "2. Select All (Ctrl+A) and Copy (Ctrl+C)"
            Write-Output "3. Paste into Notepad and save as invoice.txt"
            Write-Output "4. Run: .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.txt"
        }
        if ($ReturnObject) { return $null } else { exit 1 }
    }
}
else {
    if (-not $ReturnObject) {
        Write-Output "Usage: .\Generic-PDF-Invoice-Parser.ps1 -PDFPath <path> [-OutputFormat JSON|Config] [-ReturnObject] [-DebugMode] [-PaymentMethod Cash|Credit] [-PromptPayment] [-StrictTotals]"
        Write-Output ""
        Write-Output "Examples:"
        Write-Output "  .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.pdf"
        Write-Output "  .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.pdf -ReturnObject"
        Write-Output "  .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.pdf -PaymentMethod Credit"
    }
}
#endregion


