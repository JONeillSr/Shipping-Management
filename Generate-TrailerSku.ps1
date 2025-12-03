<#
.SYNOPSIS
    Generate standardized SKUs for JT Custom Trailers products.

.DESCRIPTION
	SKU Pattern:
        [Segment] - [Hitch] - [Family] - [Subtype?] - [Length] - [GVWR] - [Material?] - [Series?] - [Suffix?]

    Examples:
        Open bumper-pull landscape, 16 ft, 7k, steel, mid-line: OP-BP-LS-16-7K-STEEL-MID
        Hotshot low-profile gooseneck, 40 ft, 21k, commercial series: OP-GN-HS-HSLO-40-21K-COMM

    You can:
      - call New-TrailerSku directly
      - Use New-TrailerSkuFromAttributes with your attribute text
      - Bulk-process a CSV via Invoke-TrailerSkuCsv
#>

function Convert-ToTrailerCode {
    param(
        [Parameter(Mandatory)][string]$Type,
        [Parameter(Mandatory)][string]$Value
    )

    $val = $Value.Trim()

    switch ($Type.ToLower()) {
        'segment' {
            switch -regex ($val) {
                'open'      { 'OP'; break }
                'enclosed'  { 'EN'; break }
                default     { 'OP' }
            }
        }

        'hitch' {
            switch -regex ($val) {
                'bumper'       { 'BP'; break }
                'bumper pull'  { 'BP'; break }
                'gooseneck'    { 'GN'; break }
                'fifth'        { 'FW'; break }
                default        { 'BP' }
            }
        }

        'family' {
            switch -regex ($val) {
                'utility'                   { 'UT'; break }
                'landscape'                 { 'LS'; break }
                'car hauler'                { 'CH'; break }
                'equipment'                 { 'EQ'; break }
                'deckover|flatbed'          { 'DO'; break }
                'motorsports'               { 'MS'; break }
                'boat'                      { 'BT'; break }
                'dump'                      { 'DP'; break }
                'lowboy|heavy equipment'    { 'LB'; break }
                'multi-?car'                { 'MCAR'; break }
                'lift-?hauler|scissor'      { 'LH'; break }
                'hotshot'                   { 'HS'; break }
                default                      { 'UT' }
            }
        }

        'hotshotconfig' {
            switch -regex ($val) {
                'deckover'         { 'HSDO'; break }
                'low.?profile'     { 'HSLO'; break }
                'hydraulic.*dove'  { 'HSHD'; break }
                'mega.*ramp'       { 'HSMR'; break }
                'commercial'       { 'HSCM'; break }
                default            { '' }
            }
        }

        default { $val.ToUpper() }
    }
}

function New-TrailerSku {
    param(
        [Parameter(Mandatory)][ValidateSet('OP','EN')]
        [string]$Segment,

        [Parameter(Mandatory)][ValidateSet('BP','GN','FW')]
        [string]$Hitch,

        [Parameter(Mandatory)]
        [string]$FamilyCode,

        [string]$SubtypeCode,

        [Parameter(Mandatory)]
        [int]$LengthFeet,

        [Parameter(Mandatory)]
        [int]$GvwrLbs,

        [string]$Material,
        [string]$Series,
        [string]$Suffix
    )

    $lenPart   = '{0:00}' -f $LengthFeet
    $gvwrK     = [math]::Round($GvwrLbs / 1000.0)
    $gvwrPart  = '{0}K' -f $gvwrK

    $parts = @($Segment, $Hitch, $FamilyCode)
    if ($SubtypeCode) { $parts += $SubtypeCode }

    $parts += $lenPart
    $parts += $gvwrPart

    if ($Material) { $parts += $Material.ToUpper() }
    if ($Series)   { $parts += $Series.ToUpper() }
    if ($Suffix)   { $parts += $Suffix.ToUpper() }

    # Remove any accidental double dashes
    ($parts -join '-').Replace('--','-').Trim('-')
}

function New-TrailerSkuFromAttributes {
    <#
    .SYNOPSIS
        Generates a SKU using human-readable attribute values.

    .PARAMETER Segment
        'Open' or 'Enclosed'

    .PARAMETER Hitch
        'Bumper Pull', 'Gooseneck', 'Fifth-Wheel'

    .PARAMETER Family
        'Utility', 'Landscape', 'Car Hauler', 'Equipment', 'Deckover / Flatbed',
        'Motorsports', 'Boat', 'Dump', 'Lowboy / Heavy Equipment', 'Multi-Car Hauler',
        'Lift-Hauler / Scissor-Lift', 'Hotshot'

    .PARAMETER HotshotConfig
        For hotshot trailers: 'Deckover', 'Low-Profile', 'Hydraulic Dove', 'Mega-Ramp', 'Commercial Duty'

    #>
    param(
        [Parameter(Mandatory)][string]$Segment,
        [Parameter(Mandatory)][string]$Hitch,
        [Parameter(Mandatory)][string]$Family,
        [string]$HotshotConfig,
        [Parameter(Mandatory)][int]$LengthFeet,
        [Parameter(Mandatory)][int]$GvwrLbs,
        [string]$Material,
        [string]$Series,
        [string]$Suffix
    )

    $segCode   = Convert-ToTrailerCode -Type 'segment' -Value $Segment
    $hitchCode = Convert-ToTrailerCode -Type 'hitch'   -Value $Hitch
    $famCode   = Convert-ToTrailerCode -Type 'family'  -Value $Family
    $subCode   = ''

    if ($famCode -eq 'HS' -and $HotshotConfig) {
        $subCode = Convert-ToTrailerCode -Type 'hotshotconfig' -Value $HotshotConfig
    }

    New-TrailerSku -Segment $segCode -Hitch $hitchCode -FamilyCode $famCode `
        -SubtypeCode $subCode -LengthFeet $LengthFeet -GvwrLbs $GvwrLbs `
        -Material $Material -Series $Series -Suffix $Suffix
}

function Invoke-TrailerSkuCsv {
    <#
    .SYNOPSIS
        Bulk-generate SKUs from a CSV of trailer data.

    .DESCRIPTION
        Input CSV must have at least:
            Segment, Hitch, Family, LengthFt, GvwrLbs

        Optional columns:
            HotshotConfig, Material, Series, Suffix

        Adds a 'sku' column and writes a new CSV.
    #>
    param(
        [Parameter(Mandatory)][string]$InputCsvPath,
        [Parameter(Mandatory)][string]$OutputCsvPath
    )

    $items = Import-Csv -Path $InputCsvPath

    foreach ($item in $items) {
        $sku = New-TrailerSkuFromAttributes `
            -Segment      $item.Segment `
            -Hitch        $item.Hitch `
            -Family       $item.Family `
            -HotshotConfig $item.HotshotConfig `
            -LengthFeet   [int]$item.LengthFt `
            -GvwrLbs      [int]$item.GvwrLbs `
            -Material     $item.Material `
            -Series       $item.Series `
            -Suffix       $item.Suffix

        $item | Add-Member -NotePropertyName sku -NotePropertyValue $sku -Force
    }

    $items | Export-Csv -Path $OutputCsvPath -NoTypeInformation
}
