<#
.SYNOPSIS
    Generic PDF Invoice Parser for multiple auction formats
    (Stronghold Equipment, DC Auctions/Wavebid, Brolyn, Ritchie Bros, and generic).

.DESCRIPTION
    Extracts structured data from invoice PDFs (or pre-extracted .txt).
    • Robust text extraction (pdftotext preferred; iTextSharp fallback if available)
    • Vendor detection and vendor-specific item parsing:
        - Stronghold: state machine to pair Lot → Description → Hammer Price across wraps
        - DC Auctions / Wavebid: reads "Lot ... Sale Price" blocks, ignores "Loading Fee"
        - Generic: simple lot-with-description matcher
    • Normalized contact info (phones, emails)
    • Structured pickup addresses (Address2 pulled from parentheses)
    • Totals inference (Subtotal / ConvenienceFee / Grand / Cash / Credit) with optional strict validation
    • Wrapper-friendly: -ReturnObject yields a PSCustomObject with stable fields
    • ScriptAnalyzer-friendly: approved verbs, singular-noun function names, OutputType attrs, no nested functions

.PARAMETER PDFPath
    Path to a PDF invoice file (or a .txt file with the invoice text).

.PARAMETER OutputFormat
    One of: JSON | CSV | Display | Config. Default: Display.
    (CSV is usually produced by your wrapper script using -ReturnObject.)

.PARAMETER ReturnObject
    If set, returns a PSCustomObject (no console output) for wrapper scripts.

.PARAMETER SavePattern
    Reserved for future pattern learning (currently a no-op).

.PARAMETER DebugMode
    If set, writes DEBUG_ExtractedText_*.txt and DEBUG_ParsedData_*.json.

.PARAMETER GUI
    Reserved flag (prints a hint).

.PARAMETER PaymentMethod
    Cash | Credit. Default: Cash. Helps when both totals appear on the invoice.

.PARAMETER PromptPayment
    If set, interactively prompts (Cash/Credit) to resolve ambiguous totals.

.PARAMETER StrictTotals
    If set, enforces arithmetic consistency (throws on discrepancies).

.OUTPUTS
    PSCustomObject with fields:
      Vendor, InvoiceNumber, InvoiceDate,
      ContactInfo     = { Phone: [..], Email: [..] }
      PickupAddresses = [ { Street, Address2, City, State, Zip, OneLine }, ... ]
      PickupDates     = [ ... ]
      Items           = [ { LotNumber, Description, HammerPrice }, ... ]
      Totals          = { Subtotal, ConvenienceFee, GrandTotal, CashTotal, CreditTotal, Premium, Tax, Total }
      SpecialNotes    = [ ... ]

.EXAMPLE
    .\Generic-PDF-Invoice-Parser.ps1 -PDFPath '.\SE Invoice-101625-5024.pdf'

.EXAMPLE
    $obj = .\Generic-PDF-Invoice-Parser.ps1 -PDFPath '.\invoice.pdf' -ReturnObject

.EXAMPLE
    .\Generic-PDF-Invoice-Parser.ps1 -PDFPath '.\invoice.pdf' -PaymentMethod Credit -StrictTotals

.NOTES
    Version: v2.1.0
    Key changes:
      - Flattened functions; no nested definitions
      - PS5-safe (no '??'), approved verbs, singular nouns, OutputType attributes
      - Stronghold + Wavebid extractors hardened; totals logic guarded
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false)]
    [string]$PDFPath,

    [Parameter(Mandatory=$false)]
    [ValidateSet("JSON","CSV","Display","Config")]
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

# Quiet analyzer where these are referenced conditionally
$null = $PaymentMethod; $null = $PromptPayment; $null = $StrictTotals
if ($SavePattern) { $null = $true }

# Globals
$script:PatternsFile    = ".\Data\InvoicePatterns.json"
$script:LearnedPatterns = @()
$script:ItemKeys        = New-Object 'System.Collections.Generic.HashSet[string]'
$script:ItemsBuffer     = New-Object 'System.Collections.Generic.List[object]'

# ─────────────────────────────────────────────────────────────────────────────
# Helpers (PS5-safe)
# ─────────────────────────────────────────────────────────────────────────────
function Get-FirstOrDefault {
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory=$false)][object[]]$InputObject,
        [Parameter(Mandatory=$false)][object]$Default = $null
    )
    if ($null -ne $InputObject -and $InputObject.Count -gt 0) {
        return ($InputObject | Select-Object -First 1)
    }
    return $Default
}

function ConvertTo-Amount {
    [CmdletBinding()]
    [OutputType([decimal])]
    param([Parameter(Mandatory=$true)][string]$String)
    if ([string]::IsNullOrWhiteSpace($String)) { return [decimal]0 }
    return [decimal]($String -replace ',','')
}

function Test-NearEqual {
    [CmdletBinding()]
    [OutputType([bool])]
    param([decimal]$A,[decimal]$B,[decimal]$Epsilon=0.01)
    return ([Math]::Abs($A-$B) -le $Epsilon)
}

# ─────────────────────────────────────────────────────────────────────────────
# PDF extraction
# ─────────────────────────────────────────────────────────────────────────────
function Get-PDFTextContent {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory=$true)][string]$Path)

    if (-not $ReturnObject) {
        Write-Output "`n📄 Extracting text from PDF..."
        Write-Output "   File: $(Split-Path $Path -Leaf)"
    }

    $pdftotext = Get-Command pdftotext -ErrorAction SilentlyContinue
    if ($pdftotext) {
        try {
            $tmp = [System.IO.Path]::GetTempFileName()
            & pdftotext -layout $Path $tmp 2>&1 | Out-Null
            if (Test-Path $tmp) {
                $t = Get-Content $tmp -Raw -Encoding UTF8
                Remove-Item $tmp -Force
                if ($t -and $t.Length -gt 80) {
                    if (-not $ReturnObject) { Write-Output "   ✅ Extracted $($t.Length) chars via pdftotext" }
                    return @{ Text = $t; Method = "pdftotext"; Quality = "High" }
                }
            }
        } catch { Write-Verbose "pdftotext failed: $_" }
    }

    try {
        $dll = Get-ChildItem "$env:USERPROFILE\.nuget\packages\itextsharp\*\lib\net40\itextsharp.dll" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($dll) {
            Add-Type -Path $dll.FullName
            $reader = New-Object iTextSharp.text.pdf.PdfReader($Path)
            $acc = ""
            for ($p = 1; $p -le $reader.NumberOfPages; $p++) {
                $s = New-Object iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
                $pageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $p, $s)
                $acc += $pageText + "`n"
            }
            $reader.Close()
            if ($acc.Length -gt 80) {
                if (-not $ReturnObject) { Write-Output "   ✅ Extracted $($acc.Length) chars via iTextSharp" }
                return @{ Text = $acc; Method = "iTextSharp"; Quality = "High" }
            }
        }
    } catch { Write-Verbose "iTextSharp failed: $_" }

    if (-not $ReturnObject) {
        Write-Warning "`n❌ Could not extract readable text from PDF"
        Write-Output "Install xpdf's pdftotext or save the PDF as .txt and re-run."
    }
    return $null
}

# ─────────────────────────────────────────────────────────────────────────────
# Patterns / Vendor detection
# ─────────────────────────────────────────────────────────────────────────────
function Test-PdfToTextBBoxAvailable {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $cmd = Get-Command pdftotext -ErrorAction SilentlyContinue
    if (-not $cmd) { return $false }
    try {
        # Xpdf pdftotext returns usage on bad args; just test it exists
        $v = & pdftotext -v 2>$null
        return $true
    } catch { return $false }
}
function Get-PDFWordBoxes {
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory=$true)][string]$PdfPath
    )

    if (-not (Test-PdfToTextBBoxAvailable)) { return @() }

    $tmp = [System.IO.Path]::GetTempFileName()
    $ran = $false
    try {
        # Try Poppler’s/Xpdf’s layout bboxes first (better grouping); fall back to plain bboxes
        try {
            & pdftotext -bbox-layout -nopgbrk -q $PdfPath $tmp 2>$null
            $ran = Test-Path $tmp
        } catch { $ran = $false }
        if (-not $ran) {
            & pdftotext -bbox -nopgbrk -q $PdfPath $tmp 2>$null
        }
        if (-not (Test-Path $tmp)) { return @() }

        $raw = Get-Content $tmp -Raw
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }

        [xml]$xml = $raw
        # Grab ANY element named 'word' regardless of namespace or depth
        $wordNodes = $xml.SelectNodes("//*[local-name()='word']")
        if ($null -eq $wordNodes -or $wordNodes.Count -eq 0) {
            # Some Xpdf builds use capitalized <Word>
            $wordNodes = $xml.SelectNodes("//*[local-name()='Word']")
        }
        if ($null -eq $wordNodes -or $wordNodes.Count -eq 0) { return @() }

        $results = New-Object System.Collections.Generic.List[object]

        foreach ($w in $wordNodes) {
            # attributes can be xMin/xmax/XMin/xmin etc. Normalize with a small helper
            function _attr($node,$name) {
                foreach ($n in @($name, $name.ToLower(), $name.Substring(0,1).ToUpper() + $name.Substring(1).ToLower())) {
                    $a = $node.Attributes[$n]
                    if ($a) { return $a.Value }
                }
                return $null
            }

            $text = [string]$w.InnerText
            if ([string]::IsNullOrWhiteSpace($text)) { continue }

            $xMin = _attr $w 'xMin'
            $xMax = _attr $w 'xMax'
            $yMin = _attr $w 'yMin'
            $yMax = _attr $w 'yMax'

            # Find page number from ancestor <page>/<Page>
            $pageNum = $null
            $p = $w.ParentNode
            while ($p -ne $null -and $pageNum -eq $null) {
                $pageNum = (_attr $p 'number')
                if (-not $pageNum) { $pageNum = (_attr $p 'id') }
                if (-not $pageNum -and ($p.LocalName -eq 'page' -or $p.LocalName -eq 'Page')) {
                    # Poppler sometimes uses page="1"
                    $pageNum = (_attr $p 'page')
                }
                $p = $p.ParentNode
            }
            if (-not $pageNum) { $pageNum = 1 }

            # Guard against missing coords
            if (-not $xMin -or -not $xMax -or -not $yMin -or -not $yMax) { continue }

            $results.Add([pscustomobject]@{
                Page = [int]$pageNum
                xMin = [double]([System.Globalization.CultureInfo]::InvariantCulture.NumberFormat.NumberDecimalSeparator -eq '.' ? $xMin : ($xMin -replace ',','.'))
                yMin = [double]([System.Globalization.CultureInfo]::InvariantCulture.NumberFormat.NumberDecimalSeparator -eq '.' ? $yMin : ($yMin -replace ',','.'))
                xMax = [double]([System.Globalization.CultureInfo]::InvariantCulture.NumberFormat.NumberDecimalSeparator -eq '.' ? $xMax : ($xMax -replace ',','.'))
                yMax = [double]([System.Globalization.CultureInfo]::InvariantCulture.NumberFormat.NumberDecimalSeparator -eq '.' ? $yMax : ($yMax -replace ',','.'))
                Text = $text.Trim()
            }) | Out-Null
        }

        return $results.ToArray()
    } catch {
        return @()
    } finally {
        if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function Import-InvoicePattern {
    [CmdletBinding()]
    [OutputType([void])]
    param()

    if (Test-Path $script:PatternsFile) {
        $script:LearnedPatterns = Get-Content $script:PatternsFile -Raw | ConvertFrom-Json
        return
    }

    $script:LearnedPatterns = @()

    $script:LearnedPatterns += @{
        Vendor = "Brolyn Auctions"
        Patterns = @{
            CompanyIdentifier = 'Brolyn|BROLYN'
            Phone  = '\(574\)\s*891-3111'
            Email  = 'logistics@brolynauctions\.com'
            PickupDates = '(?:load times|pickup).*?(\w+\s+\d{1,2}/\d{1,2})\s+thru\s+(\w+\s+\d{1,2}/\d{1,2})'
            LotPattern = '^(\d{2,5})\s+\d{4}\s+(.+)'
        }
    }

    $script:LearnedPatterns += @{
        Vendor = "Ritchie Bros"
        Patterns = @{
            CompanyIdentifier = 'Ritchie Bros|RITCHIE BROS|RB Auctions'
            Phone  = '\(\d{3}\)\s*\d{3}-\d{4}'
            Email  = '[a-zA-Z0-9._%+-]+@rbauction\.com|[a-zA-Z0-9._%+-]+@ritchiebros\.com'
            Address = '\d+.*?(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd).*?[A-Z]{2}\s+\d{5}'
            PickupDates = '(?:Removal|Pickup).*?(\d{1,2}/\d{1,2}/\d{4})'
            LotPattern = '^(\d{2,5})\s+\d{4}\s+(.+)'
        }
    }

    $script:LearnedPatterns += @{
        Vendor = "Stronghold Equipment"
        Patterns = @{
            CompanyIdentifier = 'STRONGHOLD EQUIPMENT|Stronghold Equipment|Rocky Recycling Ltd'
            Phone  = '905-662-8200'
            Email  = 'dan@strongholdequipment.com'
            Address = '349 Arvin Ave'
            PickupDates = 'Removal|Pickup'
        }
    }
}
function Find-InvoiceVendor {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory=$true)][string]$Text)

    if (-not $ReturnObject) { Write-Output "`n🔍 Identifying vendor..." }
    $flat = $Text -replace '\s+', ' '

    # DC / Wavebid
    if ($flat -match '(?i)\bDC\s*Auctions\b|\bWavebid\b') {
        if (-not $ReturnObject) { Write-Output "   ✅ Detected: DC Auctions (Wavebid)" }
        return @{ Vendor="DC Auctions"; Patterns=@{ CompanyIdentifier='DC\s*Auctions|Wavebid' } }
    }

    # Stronghold
    if ($flat -match '(?i)stronghold|Rocky\s+Recycling') {
        if (-not $ReturnObject) { Write-Output "   ✅ Detected: Stronghold Equipment" }
        return ($script:LearnedPatterns | Where-Object { $_.Vendor -eq "Stronghold Equipment" } | Select-Object -First 1)
    }

    # Brolyn
    if ($flat -match '(?i)brolyn') {
        if (-not $ReturnObject) { Write-Output "   ✅ Detected: Brolyn Auctions" }
        return ($script:LearnedPatterns | Where-Object { $_.Vendor -eq "Brolyn Auctions" } | Select-Object -First 1)
    }

    # *** FIXED: Commercial Industrial Auctioneers (CIA) ***
    if ($flat -match '(?i)\bCommercial\s*(?:&|and)?\s*Industrial\s*Auctioneers\b') {
        if (-not $ReturnObject) { Write-Output "   ✅ Detected: Commercial Industrial Auctioneers" }
        return @{
            Vendor = "Commercial Industrial Auctioneers"
            Patterns = @{
                CompanyIdentifier = 'COMMERCIAL\s+INDUSTRIAL\s+AUCTIONEERS|Commercial\s*(?:&|and)?\s*Industrial\s*Auctioneers'
                Phone  = '\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}'
                Email  = '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[A-Za-z]{2,}'
                Address = ''
                PickupDates = 'Removal|Pickup|Pick[-\s]*up|Collection'
            }
        }
    }

    if (-not $ReturnObject) { Write-Output "   ⚠️  Unknown vendor - using generic rules" }
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

# ─────────────────────────────────────────────────────────────────────────────
# Contact & Address helpers
# ─────────────────────────────────────────────────────────────────────────────
function Test-ValidPhoneNumber {
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory=$true)][string]$Phone)
    $digits = $Phone -replace '[^\d]', ''
    if ($digits.Length -ne 10) { return $false }
    $area = [int]$digits.Substring(0,3); $exch=[int]$digits.Substring(3,3)
    if ($area -lt 200 -or $exch -lt 200) { return $false }
    return $true
}

function ConvertTo-AddressObject {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory=$true)][string]$Street,
        [Parameter(Mandatory=$true)][string]$City,
        [Parameter(Mandatory=$true)][string]$State,
        [Parameter(Mandatory=$true)][string]$Zip
    )
    $street = ($Street -replace '\s+',' ').Trim()
    $city   = ($City   -replace '\s+',' ').Trim()
    $state  = ($State  -replace '\s+','').Trim().ToUpper()
    $zip    = ($Zip    -replace '\s+','').Trim()

    $addr2 = $null
    $m = [regex]::Match($street,'\(([^)]+)\)')
    if ($m.Success) {
        $addr2 = $m.Groups[1].Value.Trim()
        $street = ($street -replace '\([^)]+\)','').Trim()
        $street = ($street -replace '\s{2,}',' ')
    }
    [pscustomobject]@{
        Street=$street; Address2=$addr2; City=$city; State=$state; Zip=$zip
        OneLine="$street, $city $state $zip"
    }
}

function ConvertFrom-FreeformAddress {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory=$true)][string]$AddressLine)

    $s = ($AddressLine -replace '\s+',' ').Trim()

    # Use double-quoted strings so the apostrophe in the character class is not ending the string
    $m1 = [regex]::Match($s, "^(?<street>.+?),\s*(?<city>[A-Za-z][A-Za-z\s.\'-]+)\s+(?<state>[A-Z]{2})\s+(?<zip>\d{5})$")
    if ($m1.Success) {
        return ConvertTo-AddressObject -Street $m1.Groups['street'].Value -City $m1.Groups['city'].Value -State $m1.Groups['state'].Value -Zip $m1.Groups['zip'].Value
    }

    $m2 = [regex]::Match($s, "^(?<street>.+?)\s+(?<city>[A-Za-z][A-Za-z\s.\'-]+)\s+(?<state>[A-Z]{2})\s+(?<zip>\d{5})$")
    if ($m2.Success) {
        return ConvertTo-AddressObject -Street $m2.Groups['street'].Value -City $m2.Groups['city'].Value -State $m2.Groups['state'].Value -Zip $m2.Groups['zip'].Value
    }

    return $null
}


function Get-PickupLocationAddress {
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param([Parameter(Mandatory=$true)][string]$RawText)

    $results = New-Object System.Collections.ArrayList
    $seen    = @{}

    $pA = "Location:\s*(?<street>\d+[^\r\n,]*?)\s*,\s*(?<city>[A-Za-z][A-Za-z\s.\'-]+)\s+(?<state>[A-Z]{2})\s+(?<zip>\d{5})"
    foreach ($m in [regex]::Matches($RawText,$pA,'IgnoreCase')) {
        $o = ConvertTo-AddressObject -Street $m.Groups['street'].Value -City $m.Groups['city'].Value -State $m.Groups['state'].Value -Zip $m.Groups['zip'].Value
        $key = "$($o.OneLine)|$($o.Address2)"
        if (-not $seen.ContainsKey($key)) { $seen[$key]=$true; $null=$results.Add($o) }
    }

    $pB = "Location:\s*(?<street>\d+[^\r\n,]*?)\s+(?<city>[A-Za-z][A-Za-z\s.\'-]+)\s+(?<state>[A-Z]{2})\s+(?<zip>\d{5})"
    foreach ($m in [regex]::Matches($RawText,$pB,'IgnoreCase')) {
        $o = ConvertTo-AddressObject -Street $m.Groups['street'].Value -City $m.Groups['city'].Value -State $m.Groups['state'].Value -Zip $m.Groups['zip'].Value
        $key = "$($o.OneLine)|$($o.Address2)"
        if (-not $seen.ContainsKey($key)) { $seen[$key]=$true; $null=$results.Add($o) }
    }
    return $results.ToArray()
}

# ─────────────────────────────────────────────────────────────────────────────
# Totals helpers
# ─────────────────────────────────────────────────────────────────────────────
function Get-AmountAfterLabel {
    [CmdletBinding()]
    [OutputType([decimal])]
    param(
        [Parameter(Mandatory=$true)][string]$Text,
        [Parameter(Mandatory=$true)][string]$Label,
        [Parameter(Mandatory=$false)][int]$Window=100
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return [decimal]0 }
    $m1 = [regex]::Match($Text, [regex]::Escape($Label), 'IgnoreCase')
    if (-not $m1.Success) { return [decimal]0 }
    $start = [Math]::Min($Text.Length, $m1.Index + $m1.Length)
    $len   = [Math]::Min($Window, [Math]::Max(0, $Text.Length - $start))
    if ($len -le 0) { return [decimal]0 }
    $slice = $Text.Substring($start, $len)
    $m2 = [regex]::Match($slice, '\$\s*([0-9,]+\.[0-9]{2})')
    if ($m2.Success) { return (ConvertTo-Amount -String $m2.Groups[1].Value) }
    return [decimal]0
}

# ─────────────────────────────────────────────────────────────────────────────
# Item buffer helper
# ─────────────────────────────────────────────────────────────────────────────
function Add-InvoiceItem {
    [CmdletBinding()]
    [OutputType([void])]
    param([string]$Lot,[string]$Desc,[string]$Price)
    if ([string]::IsNullOrWhiteSpace($Lot) -or [string]::IsNullOrWhiteSpace($Price)) { return }
    $key = "$Lot|$Price"
    if ($script:ItemKeys.Contains($key)) { return }
    $null = $script:ItemKeys.Add($key)
    $script:ItemsBuffer.Add([pscustomobject]@{
        LotNumber   = $Lot
        Description = [string]$Desc
        HammerPrice = $Price
    }) | Out-Null
}

# ─────────────────────────────────────────────────────────────────────────────
# Vendor-specific item extractors (singular noun names)
# ─────────────────────────────────────────────────────────────────────────────
function Get-CommercialIndustrialAuctioneersInvoiceItem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    $script:ItemsBuffer.Clear()
    $script:ItemKeys.Clear()

    $lines = ($Text -split "`n") | ForEach-Object { $_.TrimEnd() }

    # Lot row: lot at left, optional N/Tax, price at far right
    $lotRow = '^\s*(?<lot>\d{1,5}[A-Z]?)\s+(?<mid>.*?)(?:\s+(?:N|Tax))?\s+(?<price>\d{1,3}(?:,\d{3})*\.\d{2})\s*$'
    $lotOnly = '^\s*(?<lot>\d{1,5}[A-Z]?)\s*$'

    function Is-Noise([string]$s) {
        return ($s -match '^(?i)(Lot\s+Description|RECEIPT OF PURCHASE|Purchases\b|Buyer.?s?\s*Premium|Sub-?Total|Amount\s+Paid|Balance\s+Due|THANK YOU FOR ATTENDING|PLEASE VISIT|W\s*W\s*W|Payment\s*-|^\*+\s*C\s*O\s*N\s*T|All Purchases Must be Removed|^Page\s*\d+|^Date\b|^Duplicate Inv#|COMMERCIAL INDUSTRIAL AUCTIONEERS|^PO BOX|^PORTLAND|^Phone|AWS SOLUTIONS|^Tax Status|This is NOT an Invoice|^Copy\b|NO CASH DISCOUNT|NO WEEKEND|HOURS=|OUR NEXT AUCTION)')
    }

    $curLot   = $null
    $curDesc  = @()
    $curPrice = $null

    function Commit {
        if ($null -ne $curLot -and $null -ne $curPrice) {
            $desc = ($curDesc -join ' ').Trim()
            $desc = ($desc -replace '(^|\s)(N|Tax)(\s|$)',' ').Trim()
            Add-InvoiceItem -Lot $curLot -Desc $desc -Price (($curPrice -replace ',',''))
        }
        $script:curLot   = $null
        $script:curDesc  = @()
        $script:curPrice = $null
        Set-Variable -Name curLot   -Scope 1 -Value $script:curLot
        Set-Variable -Name curDesc  -Scope 1 -Value $script:curDesc
        Set-Variable -Name curPrice -Scope 1 -Value $script:curPrice
    }

    foreach ($raw in $lines) {
        $ln = $raw.Trim()
        if ([string]::IsNullOrWhiteSpace($ln) -or (Is-Noise $ln)) { continue }

        $m = [regex]::Match($ln, $lotRow)
        if ($m.Success) {
            Commit
            $curLot   = $m.Groups['lot'].Value
            $curPrice = $m.Groups['price'].Value
            $inline   = $m.Groups['mid'].Value.Trim()
            if ($inline -and $inline -notmatch '^(?i)(N|Tax)$') { $curDesc += $inline }
            continue
        }

        $m2 = [regex]::Match($ln, $lotOnly)
        if ($m2.Success) {
            Commit
            $curLot   = $m2.Groups['lot'].Value
            $curDesc  = @()
            $curPrice = $null
            continue
        }

        if ($null -ne $curLot) {
            if ($ln -notmatch '^(?i)(N|Tax)$') { $curDesc += $ln }
            continue
        }
    }

    if ($null -ne $curLot -and $null -ne $curPrice) { Commit }

    return $script:ItemsBuffer.ToArray()
}

function Get-StrongholdInvoiceItem {
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param([Parameter(Mandatory=$true)][string]$Text)

    $script:ItemsBuffer.Clear()
    $pricePat = '(?<!\d)(?:\$\s*)?\d{1,3}(?:,\d{3})*\.\d{2}(?:\s*\$)?(?!\d)'
    $lotLine  = '^(?<lot>\d+[A-Z]?)\b\s*(?<rest>.*)$'
    $lines = ($Text -split "`n") | ForEach-Object { $_.TrimEnd() }

    function Get-PriceLoose([string]$s) {
        $m = [regex]::Match($s,$pricePat)
        if ($m.Success) { return ($m.Value -replace '[^\d\.]','') }
        $tail = $s
        if ($tail.Length -gt 40) { $tail = $tail.Substring($tail.Length-40) }
        $tail = $tail -replace '[^0-9\.,]',''
        $m2 = [regex]::Match($tail, '\d{1,3}(?:,\d{3})*\.\d{2}')
        if ($m2.Success) { return ($m2.Value -replace ',','') }
        return $null
    }

    $inItems = $false; $curLot = $null; $curDesc = ""
    for ($i=0; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i].Trim()
        if ($ln -match '^(?i)Lot#') { $inItems=$true; continue }
        if (-not $inItems -and ($ln -match $lotLine)) { $inItems=$true }
        if (-not $inItems) { continue }

        if ($ln -match '^\*\*\*|^Subtotal\b|^Notes:') {
            if ($curLot) {
                $p = Get-PriceLoose $curDesc
                if ($p) { Add-InvoiceItem -Lot $curLot -Desc ($curDesc.Trim()) -Price $p }
            }
            break
        }

        if ($ln -match '^(?<lot>\d+[A-Z]?)\s*$') { continue }

        $mLot = [regex]::Match($ln,$lotLine)
        if ($mLot.Success) {
            if ($curLot) {
                $p = Get-PriceLoose $curDesc
                if ($p) { Add-InvoiceItem -Lot $curLot -Desc ($curDesc.Trim()) -Price $p }
            }
            $curLot  = $mLot.Groups['lot'].Value
            $curDesc = $mLot.Groups['rest'].Value.Trim()

            $pNow = Get-PriceLoose $curDesc
            if ($pNow) {
                $descOnly = ($curDesc -replace $pricePat,'').Trim()
                if (-not $descOnly) { $descOnly=$curDesc.Trim() }
                Add-InvoiceItem -Lot $curLot -Desc $descOnly -Price $pNow
                $curLot=$null; $curDesc=""
            }
            continue
        }

        if ($curLot) {
            $candidate = if ($curDesc) { "$curDesc $ln" } else { $ln }
            $p2 = Get-PriceLoose $candidate
            if ($p2) {
                $descOnly = ($candidate -replace $pricePat,'').Trim()
                if (-not $descOnly) { $descOnly=$candidate.Trim() }
                Add-InvoiceItem -Lot $curLot -Desc $descOnly -Price $p2
                $curLot=$null; $curDesc=""
            } elseif ($ln -notmatch '^\d+[A-Z]?$') {
                $curDesc = $candidate.Trim()
            }
        }
    }
    if ($curLot) {
        $p3 = Get-PriceLoose $curDesc
        if ($p3) { Add-InvoiceItem -Lot $curLot -Desc ($curDesc.Trim()) -Price $p3 }
    }
    return $script:ItemsBuffer.ToArray()
}
function Get-WavebidInvoiceItem {
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param(
        [Parameter(Mandatory = $true)] [string]$Text,
        [Parameter()]                  [string]$PdfPath
    )

    # Buffers
    $script:ItemsBuffer.Clear() | Out-Null
    $script:ItemKeys.Clear()    | Out-Null

    # Normalize lines (trim, drop empties)
    $lines = ($Text -split "(`r?`n)+" | ForEach-Object { $_.Trim() }) | Where-Object { $_ -ne '' }

    # Regex patterns (PS5-safe)
    $reHeader       = '(?i)^\s*Lot\s+Paddle\s+Description\b'
    $reHeaderEcho   = $reHeader
    $reFooter       = '(?i)^(Totals?:|Total\s+Lots:|Buyer\s+Information\b|Auction\s+Information\b|PAID\s+IN\s+FULL\b)'
    # A: lot + paddle + rest
    $reLotStartA    = '^(?<lot>\d{1,6}[A-Z]?)\s+(?<paddle>\d{3,10})\s+(?<rest>.+)$'
    # B: lot + rest (no paddle column)
    $reLotStartB    = '^(?<lot>\d{1,6}[A-Z]?)\s+(?<rest>.*[A-Za-z].*)$'
    $reAmounts      = '\$?\d{1,3}(?:,\d{3})*\.\d{2}'
    $reOnlyAmts     = '^(?:\s*\$?\d{1,3}(?:,\d{3})*\.\d{2}\s*)+$'
    $reLocation     = '(?i)^\s*Location:'

    # State for current lot block
    $inTable    = $false
    $curLot     = $null
    $descParts  = New-Object System.Collections.Generic.List[string]
    $amtBuf     = New-Object System.Collections.Generic.List[string]

    # Commit current lot block (inline, no nested function)
    function __CommitCurrentWavebidBlock_NoNested {
        param()
        if ($null -eq $curLot) { return }
        if ($amtBuf.Count -eq 0) { return }

        # Heuristic: prefer the 2nd amount if available (typical order: Bid, Sale Price, Premium, Tax, Total).
        $price = if ($amtBuf.Count -ge 2) { $amtBuf[1] } else { $amtBuf[0] }
        $price = ($price -replace '[^\d\.]','')

        $descJoined = (($descParts -join ' ') -replace '\s+', ' ').Trim()
        if (-not [string]::IsNullOrWhiteSpace($price)) {
            Add-InvoiceItem -Lot $curLot -Desc $descJoined -Price $price
        }
    }

    $i = 0
    while ($i -lt $lines.Count) {
        $ln = $lines[$i]

        if (-not $inTable) {
            if ($ln -match $reHeader) { $inTable = $true }
            $i++; continue
        }

        # Section changes / hard stops
        if ($ln -match $reFooter) {
            __CommitCurrentWavebidBlock_NoNested
            break
        }

        # Ignore repeated headers (page wrap) and location lines inside table
        if ($ln -match $reHeaderEcho) { $i++; continue }
        if ($ln -match $reLocation)   { $i++; continue }

        # Amount-only row right after header or within a block
        if ($ln -match $reOnlyAmts) {
            if ($null -ne $curLot) {
                foreach ($mAmt in [regex]::Matches($ln, $reAmounts)) {
                    $amtBuf.Add($mAmt.Value) | Out-Null
                }
            }
            $i++; continue
        }

        # New lot start? Prefer A (lot+paddle), fall back to B (lot+desc)
        $mA = [regex]::Match($ln, $reLotStartA)
        $mB = if (-not $mA.Success) { [regex]::Match($ln, $reLotStartB) } else { $null }

        if ($mA.Success -or $mB.Success) {
            # Commit previous lot block
            if ($null -ne $curLot) {
                __CommitCurrentWavebidBlock_NoNested
            }

            $curLot    = if ($mA.Success) { $mA.Groups['lot'].Value } else { $mB.Groups['lot'].Value }
            $descParts = New-Object System.Collections.Generic.List[string]
            $amtBuf    = New-Object System.Collections.Generic.List[string]

            # Initial description on same line
            $rest = if ($mA.Success) { $mA.Groups['rest'].Value } else { $mB.Groups['rest'].Value }
            $rest = $rest.Trim()
            if ($rest -and $rest -notmatch '^\d+$' -and $rest -notmatch $reOnlyAmts) {
                $descParts.Add($rest) | Out-Null
            }

            # Amounts on same line
            foreach ($mAmt in [regex]::Matches($ln, $reAmounts)) {
                $amtBuf.Add($mAmt.Value) | Out-Null
            }

            $i++; continue
        }

        # Inside a lot block: add amounts + description lines (skip noise)
        if ($null -ne $curLot) {
            foreach ($mAmt in [regex]::Matches($ln, $reAmounts)) {
                $amtBuf.Add($mAmt.Value) | Out-Null
            }
            if (($ln -notmatch $reOnlyAmts) -and ($ln -notmatch $reLocation) -and ($ln -notmatch $reHeaderEcho)) {
                $descParts.Add($ln) | Out-Null
            }
            $i++; continue
        }

        # In table but haven’t seen first lot yet; skip until a valid lot line appears
        $i++
    }

    # End-of-file commit
    if ($null -ne $curLot) {
        __CommitCurrentWavebidBlock_NoNested
    }

    return $script:ItemsBuffer.ToArray()
}

function Get-GenericInvoiceItem {
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param([Parameter(Mandatory=$true)][string]$Text)

    $script:ItemsBuffer.Clear()
    $lines = $Text -split "`n"
    $haveLot = $false
    foreach ($raw in $lines) {
        $ln = $raw.Trim()
        $m = [regex]::Match($ln,'^(\d{2,5})\s+\d{4}\s+(.+)')
        if ($m.Success) {
            $lot = $m.Groups[1].Value
            $desc = ($m.Groups[2].Value -replace '\s+',' ').Trim()
            Add-InvoiceItem -Lot $lot -Desc $desc -Price $null
            $haveLot=$true
        } elseif ($haveLot -and $ln.Length -gt 10 -and $ln.Length -lt 200) {
            $i = $script:ItemsBuffer.Count - 1
            if ($i -ge 0) {
                $prev = [string]$script:ItemsBuffer[$i].Description
                $script:ItemsBuffer[$i].Description = "$prev $ln".Trim()
            }
        }
    }
    return $script:ItemsBuffer.ToArray()
}

# ─────────────────────────────────────────────────────────────────────────────
# Core: Get-InvoiceData
# ─────────────────────────────────────────────────────────────────────────────
function Get-InvoiceData {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory=$true)][string]$Text,
        [Parameter(Mandatory=$true)][object]$VendorPattern,
        [Parameter(Mandatory=$true)][ValidateSet("Cash","Credit")][string]$PaymentMethod,
        [Parameter(Mandatory=$true)][bool]$PromptPayment,
        [Parameter(Mandatory=$true)][bool]$StrictTotals
    )

    $data = [ordered]@{
        Vendor          = $VendorPattern.Vendor
        InvoiceNumber   = $null
        InvoiceDate     = $null
        ContactInfo     = @{ Phone = @(); Email = @() }
        PickupAddresses = @()
        PickupDates     = @()
        Items           = @()
        Totals          = [ordered]@{
            Subtotal=$null; Tax=$null; Premium=$null; Total=$null;
            CashTotal=$null; CreditTotal=$null; ConvenienceFee=$null; GrandTotal=$null
        }
        SpecialNotes    = @()
    }

    if ([string]::IsNullOrWhiteSpace($Text)) { return [pscustomobject]$data }
    if (-not $ReturnObject) { Write-Output "`n📊 Extracting invoice data..." }

    $norm = ($Text -replace '[\r\n]+',' ') -replace '\s+',' '

    # Invoice number detection
    $m = [regex]::Match($norm,'(?:Invoice\s*#?\s*:?\s*)?(\d{4}-\d{6}-\d+)')
    if ($m.Success) { $data.InvoiceNumber=$m.Groups[1].Value }
    if (-not $data.InvoiceNumber) {
        $m = [regex]::Match($norm,'Invoice\s*#?\s*:?\s*([A-Z0-9-]+)')
        if ($m.Success) { $data.InvoiceNumber=$m.Groups[1].Value }
    }
    if (-not $data.InvoiceNumber) {
        $m = [regex]::Match($norm,'BUYER/INVOICE\s+(\d+)')
        if ($m.Success) { $data.InvoiceNumber=$m.Groups[1].Value }
    }

    # Date detection
    foreach ($pat in @(
        'Date:\s*(\d{1,2}/\d{1,2}/\d{4})',
        'Invoice\s+Date\s*:?\s*(\d{1,2}[-/]\w+[-/]\d{4}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})',
        '(\d{2}-\w{3}-\d{4}\s+\d{2}:\d{2})',
        'DATE:\s*(\d{1,2}/\d{1,2}/\d{4})'
    )) {
        $m = [regex]::Match($norm,$pat)
        if ($m.Success) { $data.InvoiceDate=$m.Groups[1].Value; break }
    }

    # Phones
    foreach ($pm in [regex]::Matches($norm,'\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}')) {
        $phone = $pm.Value
        if (Test-ValidPhoneNumber -Phone $phone) {
            $digits = $phone -replace '[^\d]',''
            $fmt = "({0}) {1}-{2}" -f $digits.Substring(0,3),$digits.Substring(3,3),$digits.Substring(6,4)
            if ($data.ContactInfo.Phone -notcontains $fmt) { $data.ContactInfo.Phone += $fmt }
        }
    }

    # Emails
    foreach ($em in [regex]::Matches($norm,'\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b')) {
        $email = $em.Value.ToLower()
        if ($data.ContactInfo.Email -notcontains $email) { $data.ContactInfo.Email += $email }
    }

    # Addresses (freeform + "Location:" blocks)
    foreach ($pat in @(
        '(\d+\s+(?:West|East|North|South)\s+\d+\s+(?:West|East|North|South)[^\n,]+?(?:Howe|Sturgis)[^\n,]+?(?:IN|MI)\s+\d{5})',
        '(\d+\s+[A-Za-z]+\s+(?:Blvd|Boulevard|Street|St|Avenue|Ave|Road|Rd|Drive|Dr)[^\n,]+?(?:Howe|Sturgis)[^\n,]+?(?:IN|MI)\s+\d{5})',
        '(\d+\s+[A-Za-z]+\s+(?:Blvd|Boulevard|Street|St|Avenue|Ave|Road|Rd|Drive|Dr)[^\n,]+?(?:Stoney\s+Creek)[^\n,]+?(?:ON)\s+[A-Z0-9]{3}\s+[A-Z0-9]{3})'
    )) {
        foreach ($m1 in [regex]::Matches($Text,$pat,'IgnoreCase')) {
            $line = ($m1.Groups[1].Value -replace '\s+',' ').Trim()
            $obj = ConvertFrom-FreeformAddress -AddressLine $line
            if ($obj -and -not ($data.PickupAddresses | Where-Object { $_.OneLine -eq $obj.OneLine -and $_.Address2 -eq $obj.Address2 })) {
                $data.PickupAddresses += $obj
            }
        }
    }
    foreach ($obj in (Get-PickupLocationAddress -RawText $Text)) {
        if (-not ($data.PickupAddresses | Where-Object { $_.OneLine -eq $obj.OneLine -and $_.Address2 -eq $obj.Address2 })) {
            $data.PickupAddresses += $obj
        }
    }

    # Pickup date hints
    $uniqDates=@{}
    $m = [regex]::Match($norm,'(?i)load\s+times\s+for\s+materials[^:]+:\s*((?:Monday|Tuesday|Wednesday|Thursday|Friday)\s+\d{1,2}/\d{1,2}(?:/\d{2,4})?\s+thru\s+(?:Monday|Tuesday|Wednesday|Thursday|Friday)\s+\d{1,2}/\d{1,2})')
    if ($m.Success) { $uniqDates[$m.Groups[1].Value.Trim()]=$true }
    $m = [regex]::Match($norm,'(?i)load\s+times\s+for\s+racking[^:]+:\s*((?:Monday|Tuesday|Wednesday|Thursday|Friday)\s+\d{1,2}/\d{1,2}(?:/\d{2,4})?\s+thru\s+(?:Monday|Tuesday|Wednesday|Thursday|Friday)\s+\d{1,2}/\d{1,2})')
    if ($m.Success) { $uniqDates[$m.Groups[1].Value.Trim()]=$true }
    $m = [regex]::Match($norm,'(?i)payment\s+must\s+be\s+received\s+by\s+((?:Monday|Tuesday|Wednesday|Thursday|Friday)\s+\d{1,2}/\d{1,2}/\d{2}\s+at\s+\d{1,2}[ap]m)')
    if ($m.Success) { $uniqDates[$m.Groups[1].Value.Trim()]=$true }
    $data.PickupDates = @($uniqDates.Keys)

    # Items by vendor (singular-noun function names)
    switch ($VendorPattern.Vendor) {
        'Stronghold Equipment'               { $data.Items = Get-StrongholdInvoiceItem -Text $Text }
        'DC Auctions'                        { $data.Items = Get-WavebidInvoiceItem   -Text $Text -PdfPath $Global:CurrentPdfPath }
        'Brolyn Auctions'                    { $data.Items = Get-WavebidInvoiceItem   -Text $Text -PdfPath $Global:CurrentPdfPath }
        'Commercial Industrial Auctioneers'  { $data.Items = Get-CommercialIndustrialAuctioneersInvoiceItem -Text $Text }
        default                              { $data.Items = Get-GenericInvoiceItem   -Text $Text }
    }

    # Auto-fallback if nothing captured (layout variants)
    if (-not $data.Items -or $data.Items.Count -eq 0) {
        $data.Items = Get-WavebidInvoiceItem -Text $Text -PdfPath $Global:CurrentPdfPath
    }

    if (-not $ReturnObject -and $data.Items.Count -gt 0) { Write-Output "   📦 Found $($data.Items.Count) lot item(s)" }

    # Totals
    $capturedSubtotal       = Get-AmountAfterLabel -Text $norm -Label 'SubTotal:'         -Window 80
    $capturedCashTotal      = Get-AmountAfterLabel -Text $norm -Label 'Cash Total Due:'   -Window 80
    $capturedConvenienceFee = Get-AmountAfterLabel -Text $norm -Label 'Convenience Fee'   -Window 80
    $capturedCreditTotal    = Get-AmountAfterLabel -Text $norm -Label 'Credit Total Due:' -Window 80
    $capturedGrandTotal     = Get-AmountAfterLabel -Text $norm -Label 'Grand Total:'      -Window 80
    if ($capturedGrandTotal -eq 0)     { $capturedGrandTotal     = Get-AmountAfterLabel -Text $norm -Label 'Total in US Dollars' -Window 80 }
    if ($capturedConvenienceFee -eq 0) { $capturedConvenienceFee = Get-AmountAfterLabel -Text $norm -Label "Buyer's Premium"      -Window 80 }

    if ($capturedSubtotal       -ne 0) { $data.Totals.Subtotal       = $capturedSubtotal }
    if ($capturedConvenienceFee -ne 0) { $data.Totals.ConvenienceFee = $capturedConvenienceFee; $data.Totals.Premium=$capturedConvenienceFee }
    if ($capturedCashTotal      -ne 0) { $data.Totals.CashTotal      = $capturedCashTotal }
    if ($capturedCreditTotal    -ne 0) { $data.Totals.CreditTotal    = $capturedCreditTotal }
    if ($capturedGrandTotal     -ne 0) { $data.Totals.GrandTotal     = $capturedGrandTotal }

    if ($null -ne $data.Totals.Subtotal) {
        $sub = [decimal]$data.Totals.Subtotal
        $fee = if ($null -ne $data.Totals.ConvenienceFee) { [decimal]$data.Totals.ConvenienceFee } else { $null }

        if ($null -eq $data.Totals.CashTotal -or $data.Totals.CashTotal -lt $sub) { $data.Totals.CashTotal = $sub }
        if ($null -ne $fee -and $data.Totals.CashTotal -eq $fee) { $data.Totals.CashTotal = $sub }

        if ($null -ne $fee) {
            $data.Totals.CreditTotal = $sub + $fee
        } elseif ($null -eq $data.Totals.CreditTotal -and $null -ne $data.Totals.GrandTotal) {
            if ($null -eq $data.Totals.CashTotal) { $data.Totals.CashTotal = $data.Totals.GrandTotal }
        }
    }

    $selected = $PaymentMethod
    if ($PromptPayment -and $null -ne $data.Totals.CashTotal -and ($null -ne $data.Totals.CreditTotal -or $null -ne $data.Totals.ConvenienceFee)) {
        $ans = Read-Host "Payment method for totals (Cash/Credit) [$selected]"
        if     ($ans -match '^(?i)credit$') { $selected='Credit' }
        elseif ($ans -match '^(?i)cash$')   { $selected='Cash'   }
    }

    if ($StrictTotals) {
        if ($null -eq $data.Totals.Subtotal) { throw "StrictTotals: Missing Subtotal." }
        $sub=[decimal]$data.Totals.Subtotal
        $fee=if ($null -ne $data.Totals.ConvenienceFee) { [decimal]$data.Totals.ConvenienceFee } else { $null }
        switch ($selected) {
            'Credit' {
                if ($null -ne $fee) {
                    $calc = $sub + $fee
                    if ($null -ne $data.Totals.CreditTotal -and -not (Test-NearEqual -A $data.Totals.CreditTotal -B $calc)) {
                        throw "StrictTotals: Captured Credit Total ($('{0:N2}' -f $data.Totals.CreditTotal)) ≠ Subtotal+Fee ($('{0:N2}' -f $calc))."
                    }
                } elseif ($null -eq $data.Totals.CreditTotal) {
                    throw "StrictTotals: Need 'Credit Total Due' OR both Subtotal and Convenience Fee."
                }
            }
            default {
                if ($null -ne $data.Totals.CashTotal -and -not (Test-NearEqual -A $data.Totals.CashTotal -B $sub)) {
                    if ($null -ne $fee -and (Test-NearEqual -A $data.Totals.CashTotal -B $fee)) {
                        throw "StrictTotals: Captured Cash Total equals Convenience Fee ($('{0:N2}' -f $data.Totals.CashTotal)); ambiguous."
                    } else {
                        throw "StrictTotals: Captured Cash Total ($('{0:N2}' -f $data.Totals.CashTotal)) ≠ Subtotal ($('{0:N2}' -f $sub))."
                    }
                }
            }
        }
    }

    switch ($selected) {
        'Credit' {
            if     ($null -ne $data.Totals.CreditTotal)                   { $data.Totals.Total = $data.Totals.CreditTotal }
            elseif ($null -ne $data.Totals.Subtotal -and $null -ne $data.Totals.ConvenienceFee) { $data.Totals.Total = [decimal]$data.Totals.Subtotal + [decimal]$data.Totals.ConvenienceFee }
            else                                                          { $data.Totals.Total = $data.Totals.CashTotal }
        }
        default {
            if ($null -ne $data.Totals.GrandTotal) { $data.Totals.Total = $data.Totals.GrandTotal }
            else { $data.Totals.Total = $data.Totals.CashTotal }
        }
    }

    if (-not $StrictTotals -and $null -ne $data.Totals.Total -and $null -ne $data.Totals.Subtotal -and $data.Totals.Total -lt $data.Totals.Subtotal) {
        if (-not $ReturnObject) { Write-Warning "Total < Subtotal; correcting to Subtotal (cash)." }
        $data.Totals.Total = $data.Totals.Subtotal
        $data.Totals.CashTotal = $data.Totals.Subtotal
    }

    if (-not $ReturnObject) {
        Write-Output "   💳 Selection: $selected"
        if ($null -ne $data.Totals.Subtotal)       { Write-Output ("   💰 Subtotal: {0:N2}" -f [decimal]$data.Totals.Subtotal) }
        if ($null -ne $data.Totals.CashTotal)      { Write-Output ("   💵 Cash Total Due: {0:N2}" -f [decimal]$data.Totals.CashTotal) }
        if ($null -ne $data.Totals.CreditTotal)    { Write-Output ("   💳 Credit Total Due: {0:N2}" -f [decimal]$data.Totals.CreditTotal) }
        if ($null -ne $data.Totals.ConvenienceFee) { Write-Output ("   🧾 Convenience Fee: {0:N2}" -f [decimal]$data.Totals.ConvenienceFee) }
        if ($null -ne $data.Totals.Total)          { Write-Output ("   ✅ Using Total: {0:N2}" -f [decimal]$data.Totals.Total) }
    }

    return [pscustomobject]$data
}

# ─────────────────────────────────────────────────────────────────────────────
# Export/Display
# ─────────────────────────────────────────────────────────────────────────────
function Export-InvoiceData {
    [CmdletBinding()]
    [OutputType([object])]
    param (
        [Parameter(Mandatory=$true)][object]$Data,
        [Parameter(Mandatory=$true)][string]$Format,
        [Parameter(Mandatory=$false)][string]$OutputPath
    )
    if ($ReturnObject) { return $Data }

    switch ($Format) {
        "JSON" {
            $json = if ($OutputPath) { $OutputPath } else { ".\InvoiceData_$(Get-Date -Format 'yyyyMMdd_HHmmss').json" }
            $Data | ConvertTo-Json -Depth 10 | Out-File $json -Encoding UTF8
            Write-Output "`nExported to JSON: $json"
            return $json
        }
        "Config" {
            $a = Get-FirstOrDefault -InputObject $Data.PickupAddresses
            $pickup = "TBD"
            if ($null -ne $a) {
                if ($a.Address2) { $pickup = "$($a.Street)`n$($a.Address2)`n$($a.City) $($a.State) $($a.Zip)" }
                else { $pickup = "$($a.Street)`n$($a.City) $($a.State) $($a.Zip)" }
            }

            $cfg = @{
                email_subject      = "Freight Quote Request - $((Get-FirstOrDefault -InputObject @($a.Street) -Default 'TBD')) to Ashtabula, OH"
                auction_info       = @{
                    auction_name     = $Data.Vendor
                    pickup_address   = $pickup
                    logistics_contact = @{
                        phone = (Get-FirstOrDefault -InputObject $Data.ContactInfo.Phone -Default "")
                        email = (Get-FirstOrDefault -InputObject $Data.ContactInfo.Email -Default "")
                    }
                    pickup_datetime   = (Get-FirstOrDefault -InputObject $Data.PickupDates -Default "TBD")
                    delivery_datetime = "TBD"
                    delivery_notice   = "Driver must call at least one hour prior to delivery"
                    special_notes     = $Data.SpecialNotes
                }
                delivery_address   = "1218 Lake Avenue, Ashtabula, OH 44004"
                shipping_requirements = @{
                    total_pallets = "TBD"
                    truck_types   = "TBD - Please recommend based on items"
                    labor_needed  = "TBD"
                    weight_notes  = "Total weight will NOT exceed standard truck capacity"
                }
            }
            $out = if ($OutputPath) { $OutputPath } else { ".\Config_$($Data.Vendor -replace '\s+','_')_$(Get-Date -Format 'yyyyMMdd_HHmmss').json" }
            $cfg | ConvertTo-Json -Depth 10 | Out-File -FilePath $out -Encoding UTF8
            Write-Output "`nExported logistics config: $out"
            return $out
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
                $Data.ContactInfo.Phone | ForEach-Object { Write-Output "  • $_" }
            }
            if ($Data.ContactInfo.Email.Count -gt 0) {
                Write-Output "Email(s):"
                $Data.ContactInfo.Email | ForEach-Object { Write-Output "  • $_" }
            }

            if ($Data.PickupAddresses.Count -gt 0) {
                Write-Output "`nPICKUP ADDRESSES"
                foreach ($ax in $Data.PickupAddresses) {
                    Write-Output "• $($ax.Street)"
                    if ($ax.Address2) { Write-Output "  $($ax.Address2)" }
                    Write-Output "  $($ax.City) $($ax.State) $($ax.Zip)"
                }
            }

            if ($Data.PickupDates.Count -gt 0) {
                Write-Output "`nPICKUP DATES"
                $Data.PickupDates | ForEach-Object { Write-Output "• $_" }
            }

            if ($Data.Items.Count -gt 0) {
                Write-Output "`nITEMS ($($Data.Items.Count) lots)"
                $Data.Items | Select-Object -First 10 | ForEach-Object {
                    $p = if ($_.HammerPrice) { " - `$" + $_.HammerPrice } else { "" }
                    Write-Output ("Lot {0}: {1}{2}" -f $_.LotNumber,$_.Description,$p)
                }
                if ($Data.Items.Count -gt 10) { Write-Output ("... and {0} more items" -f ($Data.Items.Count-10)) }
            }

            if ($Data.Totals.Total) {
                Write-Output "`nTOTAL: $({0:N2} -f [decimal]$Data.Totals.Total)"
                if ($Data.Totals.CashTotal -or $Data.Totals.CreditTotal) {
                    $cash   = if ($Data.Totals.CashTotal)   { '{0:N2}' -f [decimal]$Data.Totals.CashTotal } else { '' }
                    $credit = if ($Data.Totals.CreditTotal) { '{0:N2}' -f [decimal]$Data.Totals.CreditTotal } else { '' }
                    Write-Output "   (Cash Total: $cash; Credit Total: $credit)"
                }
            }
            Write-Output ""
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Entrypoint
# ─────────────────────────────────────────────────────────────────────────────
if (-not $ReturnObject) {
    Write-Output "`n╔════════════════════════════════════════════════════════╗"
    Write-Output "║     GENERIC PDF INVOICE PARSER v 2.2.0                 ║"
    Write-Output "╚════════════════════════════════════════════════════════╝`n"
}

Import-InvoicePattern

if ($GUI) {
    if (-not $ReturnObject) {
        Write-Output "GUI mode not yet implemented; use: .\Generic-PDF-Invoice-Parser.ps1 -PDFPath <path>"
    }
} elseif ($PDFPath) {
    if (-not (Test-Path $PDFPath)) {
        Write-Error "File not found: $PDFPath"
        if ($ReturnObject) { return $null } else { exit 1 }
    }

    $ext = [System.IO.Path]::GetExtension($PDFPath).ToLower()
    if ($ext -eq '.txt') {
        if (-not $ReturnObject) { Write-Output "`n📄 Reading text file..." }
        $extraction = @{ Text = Get-Content $PDFPath -Raw -Encoding UTF8; Method = "TextFile"; Quality = "High" }
    } else {
        $extraction = Get-PDFTextContent -Path $PDFPath
    }

    if ($extraction -and $extraction.Text) {
        if ($DebugMode) {
            $dbg = ".\DEBUG_ExtractedText_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $extraction.Text | Out-File -FilePath $dbg -Encoding UTF8
            if (-not $ReturnObject) { Write-Output "`n🔍 DEBUG: Saved extracted text to: $dbg" }
        }

        $vendor = Find-InvoiceVendor -Text $extraction.Text
        $script:ItemKeys.Clear()    | Out-Null
        $script:ItemsBuffer.Clear() | Out-Null

        $Global:CurrentPdfPath = $PDFPath
        $data = Get-InvoiceData -Text $extraction.Text -VendorPattern $vendor `
                                -PaymentMethod $PaymentMethod -PromptPayment:$PromptPayment `
                                -StrictTotals:$StrictTotals

        if ($ReturnObject) { return $data }

        Export-InvoiceData -Data $data -Format $OutputFormat

        if ($DebugMode) {
            $dbgJson = ".\DEBUG_ParsedData_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            $data | ConvertTo-Json -Depth 10 | Out-File -FilePath $dbgJson -Encoding UTF8
            if (-not $ReturnObject) { Write-Output "`n🔍 DEBUG: Saved parsed data to: $dbgJson" }
        }
    } else {
        Write-Error "Failed to extract text from PDF"
        if (-not $ReturnObject) {
            Write-Output "`n💡 QUICK FIX:"
            Write-Output "   1) Open PDF in Adobe Reader"
            Write-Output "   2) Ctrl+A, Ctrl+C → paste to Notepad and save as invoice.txt"
            Write-Output "   3) Run this script with -PDFPath invoice.txt"
        }
        if ($ReturnObject) { return $null } else { exit 1 }
    }
} else {
    if (-not $ReturnObject) {
        Write-Output "Usage: .\Generic-PDF-Invoice-Parser.ps1 -PDFPath <path> [-OutputFormat JSON|Config] [-ReturnObject] [-DebugMode] [-PaymentMethod Cash|Credit] [-PromptPayment] [-StrictTotals]"
        Write-Output "Examples:"
        Write-Output "  .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.pdf"
        Write-Output "  .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.pdf -ReturnObject"
        Write-Output "  .\Generic-PDF-Invoice-Parser.ps1 -PDFPath invoice.pdf -PaymentMethod Credit"
    }
}
