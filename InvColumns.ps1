<#
================================================================================
 InvColumns.ps1
 The column catalogue and the terminal picker.

 EVERY COLUMN CARRIES ITS OWN RULE
   A column is not just a header and a value. Each one states the type it must
   be, the rule it must satisfy, and how it is normalised before it is written.
   A value that fails its rule is written BLANK with the reason in the Notes
   column - never a plausible wrong number in a tax workpaper.

 ADDING A COLUMN
   Add one entry here. Nothing else changes: the picker, the writer, the CSV
   and the review screen all read this list.
================================================================================
#>

Set-StrictMode -Version 2.0

function New-InvColumn {
    param(
        [string] $Key,
        [string] $Header,
        [string] $Type      = 'text',      # text | money | date | pct | int
        [int]    $Width     = 14,
        [string] $State     = 'off',       # locked | on | off
        [string] $Rule      = '',
        [scriptblock] $Normalize = $null
    )
    [pscustomobject]@{
        Key = $Key; Header = $Header; Type = $Type; Width = $Width
        State = $State; Rule = $Rule; Normalize = $Normalize
    }
}

# The first three are LOCKED: they identify the row and can never be dropped.
# The rest follow the working-paper header order, then the extras.
$InvColumnCatalog = @(

    New-InvColumn -Key 'VenInvoice' -Header 'VenInvoice' -Width 34 -State 'locked' `
        -Rule 'Vendor and invoice number joined. Blank if either is missing.'

    New-InvColumn -Key 'Source' -Header 'Source' -Width 26 -State 'locked' `
        -Rule 'The PDF file name, hyperlinked to the file itself.'

    New-InvColumn -Key 'Provided' -Header 'Provided?' -Width 10 -State 'locked' `
        -Rule 'Y only when readable text came out of the PDF. A scan reads N.'

    New-InvColumn -Key 'City' -Header 'City' -Width 18 -State 'on' `
        -Rule 'Vendor city, from the vendor address block. Title case.' `
        -Normalize { param($v) if ($v) { (Get-Culture).TextInfo.ToTitleCase(([string]$v).ToLowerInvariant()) } else { $v } }

    New-InvColumn -Key 'State' -Header 'State' -Width 8 -State 'on' `
        -Rule 'Vendor state. Must be a real two-letter US state code or it is blanked.' `
        -Normalize { param($v) if ($v) { ([string]$v).ToUpperInvariant() } else { $v } }

    New-InvColumn -Key 'Vendor' -Header 'Vendor' -Width 30 -State 'on' `
        -Rule 'Who billed us. Can never be our own company - see the identity guard.' `
        -Normalize { param($v) if ($v) { ([string]$v).Trim(' ', ',', '.', ':') } else { $v } }

    New-InvColumn -Key 'Description' -Header 'Description' -Width 46 -State 'on' `
        -Rule 'What was bought. A summary, not a parsed line-item table. Max 200 chars.'

    New-InvColumn -Key 'InvNum' -Header 'Inv#' -Width 18 -State 'on' `
        -Rule 'Must contain a digit, must not be a date or an amount, 2-32 characters. Written as text so leading zeros survive.' `
        -Normalize { param($v) if ($v) { ([string]$v).Trim() } else { $v } }

    New-InvColumn -Key 'InvDate' -Header 'Inv Date' -Type 'date' -Width 12 -State 'on' `
        -Rule 'A real date between 1995 and 30 days ahead. Written as a true date, not text.'

    New-InvColumn -Key 'TaxRate' -Header 'Tax Rate' -Type 'pct' -Width 11 -State 'on' `
        -Rule 'Stated rate, else tax divided by taxable. Must be 0-25%. Flagged if it disagrees with the amounts by more than 0.15 points.'

    New-InvColumn -Key 'InvAmt' -Header 'Inv Amt' -Type 'money' -Width 14 -State 'on' `
        -Rule 'The grand total. Checked against subtotal + freight - discount + tax.'

    New-InvColumn -Key 'Taxable' -Header 'Taxable Amount' -Type 'money' -Width 16 -State 'on' `
        -Rule 'Stated taxable amount, else derived from tax and rate, else the subtotal. Must not exceed the total.'

    New-InvColumn -Key 'TaxAmt' -Header 'Tax Amt' -Type 'money' -Width 13 -State 'on' `
        -Rule 'Sales or use tax charged. Part of the totals check.'

    # ---- beyond the working-paper header ----

    New-InvColumn -Key 'VendorAddress' -Header 'Vendor Address' -Width 34 -State 'on' `
        -Rule 'The vendor street address, as printed.'

    New-InvColumn -Key 'Subtotal' -Header 'Subtotal' -Type 'money' -Width 13 -State 'on' `
        -Rule 'Pre-tax amount. Derived from total minus tax and freight when not stated.'

    New-InvColumn -Key 'Freight' -Header 'Freight' -Type 'money' -Width 12 -State 'on' `
        -Rule 'Shipping, handling, delivery or postage. Blank means none was charged - it does not mean nothing shipped.'

    New-InvColumn -Key 'ShipToCity' -Header 'Ship To City' -Width 18 -State 'on' `
        -Rule 'Where the goods went. For use tax this is often the jurisdiction that matters, not the billing city.'

    New-InvColumn -Key 'ShipToState' -Header 'Ship To State' -Width 12 -State 'on' `
        -Rule 'Must be a real two-letter US state code or it is blanked.' `
        -Normalize { param($v) if ($v) { ([string]$v).ToUpperInvariant() } else { $v } }

    New-InvColumn -Key 'ShipToAddress' -Header 'Ship To Address' -Width 34 -State 'off' `
        -Rule 'The full ship-to address, when the invoice carries one.'

    New-InvColumn -Key 'ShipToName' -Header 'Ship To Name' -Width 26 -State 'off' `
        -Rule 'Who received the goods. Usually one of our own sites.'

    New-InvColumn -Key 'Shipped' -Header 'Shipped?' -Width 10 -State 'on' `
        -Rule 'Y when the invoice shows a ship-to block or a freight charge. Goods often ship with neither, so N is not proof nothing shipped.'

    New-InvColumn -Key 'Zip' -Header 'Zip' -Width 11 -State 'off' `
        -Rule 'Vendor postal code, 5 or 9 digit.'

    New-InvColumn -Key 'DueDate' -Header 'Due Date' -Type 'date' -Width 12 -State 'off' `
        -Rule 'Flagged when it falls before the invoice date.'

    New-InvColumn -Key 'PONum' -Header 'PO#' -Width 16 -State 'off' `
        -Rule 'Purchase order reference. Same format rule as the invoice number.'

    New-InvColumn -Key 'Discount' -Header 'Discount' -Type 'money' -Width 12 -State 'off' `
        -Rule 'Taken as a positive amount and subtracted in the totals check.'

    New-InvColumn -Key 'Terms' -Header 'Terms' -Width 14 -State 'off' `
        -Rule 'Payment terms as printed, 2-40 characters.'

    New-InvColumn -Key 'AcctNum' -Header 'Account#' -Width 16 -State 'off' `
        -Rule 'Our account number with the vendor.'

    New-InvColumn -Key 'Currency' -Header 'Currency' -Width 10 -State 'off' `
        -Rule 'USD unless the invoice states otherwise.'

    New-InvColumn -Key 'BillToName' -Header 'Bill To Name' -Width 26 -State 'off' `
        -Rule 'Who was billed. Should be us - if it is not, the invoice may be misfiled.'

    New-InvColumn -Key 'PageCount' -Header 'Pages' -Type 'int' -Width 8 -State 'off' `
        -Rule 'Pages in the PDF.'

    New-InvColumn -Key 'Confidence' -Header 'Confidence' -Type 'int' -Width 11 -State 'on' `
        -Rule 'The lowest confidence among vendor, invoice number, date, total and tax.'

    New-InvColumn -Key 'NeedsReview' -Header 'Needs Review' -Width 13 -State 'on' `
        -Rule 'Y when confidence is below 75 or a rule failed.'

    New-InvColumn -Key 'Notes' -Header 'Extraction Notes' -Width 60 -State 'on' `
        -Rule 'Every rule that fired: what was checked, derived, cleared or disagreed.'
)

function Get-InvColumnByKey {
    param([string]$Key)
    foreach ($c in $InvColumnCatalog) { if ($c.Key -eq $Key) { return $c } }
    return $null
}

function Get-InvDefaultColumnKeys {
    $keys = [System.Collections.ArrayList]::new()
    foreach ($c in $InvColumnCatalog) {
        if ($c.State -eq 'locked' -or $c.State -eq 'on') { [void]$keys.Add($c.Key) }
    }
    return @($keys)
}

function Get-InvLockedColumnKeys {
    $keys = [System.Collections.ArrayList]::new()
    foreach ($c in $InvColumnCatalog) { if ($c.State -eq 'locked') { [void]$keys.Add($c.Key) } }
    return @($keys)
}

<#
 Turns a list of keys into the column definitions to write, in catalogue order,
 with the locked columns always present and always first.
#>
function Resolve-InvColumns {
    param([AllowEmptyCollection()] $Keys)

    $want = @{}
    foreach ($k in @($Keys)) { if ($k) { $want[[string]$k] = $true } }
    foreach ($k in (Get-InvLockedColumnKeys)) { $want[$k] = $true }

    $out = [System.Collections.ArrayList]::new()
    foreach ($c in $InvColumnCatalog) {
        if ($want.ContainsKey($c.Key)) { [void]$out.Add($c) }
    }
    return @($out)
}

# ------------------------------------------------------------------- presets

function Get-InvPresetPath {
    param([string]$Root)
    if ([string]::IsNullOrEmpty($Root)) {
        $Root = $PSScriptRoot
        if ([string]::IsNullOrEmpty($Root)) { $Root = (Get-Location).Path }
    }
    return (Join-Path $Root 'config/columns.presets.json')
}

function Get-InvPresets {
    param([string]$Root)
    $p = Get-InvPresetPath -Root $Root
    if (-not (Test-Path -LiteralPath $p)) { return @{} }
    try {
        $raw = Get-Content -LiteralPath $p -Raw | ConvertFrom-Json
        $h = @{}
        foreach ($prop in $raw.PSObject.Properties) { $h[$prop.Name] = @($prop.Value) }
        return $h
    }
    catch { return @{} }
}

function Save-InvPreset {
    param([string]$Root, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)]$Keys)
    $p = Get-InvPresetPath -Root $Root
    $dir = Split-Path -Parent $p
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $all = Get-InvPresets -Root $Root
    $all[$Name] = @($Keys)

    $o = [ordered]@{}
    foreach ($k in ($all.Keys | Sort-Object)) { $o[$k] = @($all[$k]) }
    ($o | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $p -Encoding UTF8
    return $p
}

# -------------------------------------------------------------- the picker UI

<#
 The column picker.

 Numbers toggle. Ranges work (16,19-21). Locked columns show [L] and cannot be
 turned off. Presets save and load. Enter runs.

 Returns the chosen keys, or $null if the user quits.
#>
function Show-InvColumnPicker {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()] $Selected = $null,
        [string] $Root,
        [string] $PresetName = 'default'
    )

    $cat = @($InvColumnCatalog)
    $locked = @{}
    foreach ($k in (Get-InvLockedColumnKeys)) { $locked[$k] = $true }

    $on = @{}
    if ($null -eq $Selected -or @($Selected).Count -eq 0) {
        foreach ($k in (Get-InvDefaultColumnKeys)) { $on[$k] = $true }
    }
    else {
        foreach ($k in @($Selected)) { $on[[string]$k] = $true }
    }
    foreach ($k in $locked.Keys) { $on[$k] = $true }

    $presetLabel = $PresetName

    while ($true) {
        Clear-Host
        Write-Host ''
        Write-Host '  COLUMNS' -ForegroundColor Cyan -NoNewline
        Write-Host ("      preset: " + $presetLabel) -ForegroundColor DarkGray
        Write-Host '  ---------------------------------------------------------------------------' -ForegroundColor DarkGray

        # three columns of entries, numbered down each column
        $n = $cat.Count
        $rows = [int][Math]::Ceiling($n / 3.0)
        for ($r = 0; $r -lt $rows; $r++) {
            $line = '  '
            for ($c = 0; $c -lt 3; $c++) {
                $i = $r + ($c * $rows)
                if ($i -ge $n) { continue }
                $col = $cat[$i]
                $mark = '[ ]'
                if ($locked.ContainsKey($col.Key)) { $mark = '[L]' }
                elseif ($on.ContainsKey($col.Key)) { $mark = '[x]' }
                $line += ('{0,3} {1} {2,-20}' -f ($i + 1), $mark, $col.Header)
            }
            Write-Host $line
        }

        $count = 0
        foreach ($c in $cat) { if ($on.ContainsKey($c.Key)) { $count++ } }

        Write-Host '  ---------------------------------------------------------------------------' -ForegroundColor DarkGray
        Write-Host ("  [L] always written                             {0} of {1} selected" -f $count, $n) -ForegroundColor DarkGray
        Write-Host '  numbers toggle (e.g. 16,19-21)   a all   n none   d defaults   ? rule' -ForegroundColor DarkGray
        Write-Host '  p load preset    s save preset    ENTER run    q quit' -ForegroundColor DarkGray
        Write-Host ''
        Write-Host '  > ' -NoNewline -ForegroundColor Cyan

        $input = Read-Host
        if ($null -eq $input) { $input = '' }
        $t = $input.Trim()

        if ($t -eq '') {
            $keys = [System.Collections.ArrayList]::new()
            foreach ($c in $cat) { if ($on.ContainsKey($c.Key)) { [void]$keys.Add($c.Key) } }
            return @($keys)
        }

        switch -Regex ($t) {
            '^(q|quit)$' { return $null }

            '^a$' { foreach ($c in $cat) { $on[$c.Key] = $true }; continue }

            '^n$' {
                $on = @{}
                foreach ($k in $locked.Keys) { $on[$k] = $true }
                continue
            }

            '^d$' {
                $on = @{}
                foreach ($k in (Get-InvDefaultColumnKeys)) { $on[$k] = $true }
                $presetLabel = 'default'
                continue
            }

            '^\?\s*(\d+)$' {
                $i = [int]$Matches[1] - 1
                if ($i -ge 0 -and $i -lt $n) {
                    Write-Host ''
                    Write-Host ('  ' + $cat[$i].Header) -ForegroundColor Cyan
                    Write-Host ('  type: ' + $cat[$i].Type)
                    Write-Host ('  rule: ' + $cat[$i].Rule) -ForegroundColor Gray
                    Write-Host ''
                    Write-Host '  press Enter' -ForegroundColor DarkGray -NoNewline
                    [void](Read-Host)
                }
                continue
            }

            '^\?$' {
                Write-Host ''
                foreach ($c in $cat) {
                    Write-Host ('  {0,-18} {1}' -f $c.Header, $c.Rule) -ForegroundColor Gray
                }
                Write-Host ''
                Write-Host '  press Enter' -ForegroundColor DarkGray -NoNewline
                [void](Read-Host)
                continue
            }

            '^p$' {
                $presets = Get-InvPresets -Root $Root
                if ($presets.Count -eq 0) {
                    Write-Host '  no presets saved yet' -ForegroundColor Yellow
                    Start-Sleep -Milliseconds 900
                    continue
                }
                Write-Host ''
                $names = @($presets.Keys | Sort-Object)
                for ($i = 0; $i -lt $names.Count; $i++) { Write-Host ('  {0}) {1}' -f ($i + 1), $names[$i]) }
                Write-Host '  which? ' -NoNewline
                $pick = Read-Host
                $pi = 0
                if ([int]::TryParse($pick, [ref]$pi) -and $pi -ge 1 -and $pi -le $names.Count) {
                    $on = @{}
                    foreach ($k in @($presets[$names[$pi - 1]])) { $on[[string]$k] = $true }
                    foreach ($k in $locked.Keys) { $on[$k] = $true }
                    $presetLabel = $names[$pi - 1]
                }
                continue
            }

            '^s$' {
                Write-Host '  name for this preset: ' -NoNewline
                $nm = Read-Host
                if (-not [string]::IsNullOrWhiteSpace($nm)) {
                    $keys = [System.Collections.ArrayList]::new()
                    foreach ($c in $cat) { if ($on.ContainsKey($c.Key)) { [void]$keys.Add($c.Key) } }
                    [void](Save-InvPreset -Root $Root -Name $nm.Trim() -Keys $keys)
                    $presetLabel = $nm.Trim()
                    Write-Host '  saved' -ForegroundColor Green
                    Start-Sleep -Milliseconds 600
                }
                continue
            }

            default {
                foreach ($piece in ($t -split '[,\s]+')) {
                    if ($piece -eq '') { continue }
                    if ($piece -match '^(\d+)-(\d+)$') {
                        $a = [int]$Matches[1]; $b = [int]$Matches[2]
                        if ($a -gt $b) { $tmp = $a; $a = $b; $b = $tmp }
                        for ($i = $a; $i -le $b; $i++) { Switch-InvColumn -Index $i -Cat $cat -On $on -Locked $locked }
                    }
                    elseif ($piece -match '^\d+$') {
                        Switch-InvColumn -Index ([int]$piece) -Cat $cat -On $on -Locked $locked
                    }
                }
            }
        }
    }
}

function Switch-InvColumn {
    param([int]$Index, $Cat, $On, $Locked)
    $i = $Index - 1
    if ($i -lt 0 -or $i -ge @($Cat).Count) { return }
    $key = $Cat[$i].Key
    if ($Locked.ContainsKey($key)) { return }
    if ($On.ContainsKey($key)) { $On.Remove($key) } else { $On[$key] = $true }
}
