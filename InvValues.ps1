<#
================================================================================
 InvValues.ps1
 Turning a piece of text into a typed value - or refusing to.

 THE RULE THAT MATTERS
   Every parser here returns $null when the text is not unambiguously the thing
   asked for. A blank cell with a note is worth far more than a number that
   looks right and is not. Nothing here guesses.

 Pure functions, no I/O. Every one is covered by tests/Test-Values.ps1.
================================================================================
#>

Set-StrictMode -Version 2.0

# Money as printed on invoices:  1,234.56  $1,234.56  (1,234.56)  1234.56-
# 1.234,56 (European)  850.00 CR   USD 42.00
$script:InvRxMoneyUS = [regex]'^\(?\s*(?:USD|US\$|\$|EUR|GBP|CAD|£|€)?\s*(?<n>\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)\s*(?:USD|CAD|EUR|GBP)?\s*\)?\s*(?<sfx>CR|DR|-)?$'
$script:InvRxMoneyEU = [regex]'^\(?\s*(?:EUR|€|£)?\s*(?<n>\d{1,3}(?:\.\d{3})+,\d{1,2})\s*\)?$'
$script:InvRxPercent = [regex]'^\(?\s*(?<n>\d{1,3}(?:\.\d{1,4})?)\s*%\s*\)?$'

<#
 Parses money. Returns [double] or $null.

 Refuses anything that is not plainly a monetary amount: percentages, dates,
 phone numbers, ranges, and bare integers longer than 9 digits (those are
 invoice or account numbers, not dollars).

 Parentheses, a trailing minus and a trailing CR all mean negative - credits
 show up on invoices constantly and must not be read as positive.
#>
function Convert-InvMoney {
    [CmdletBinding()]
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = $Text.Trim()

    # obvious non-money shapes
    if ($t -match '%') { return $null }
    if ($t -match '\d[/]\d') { return $null }                    # 03/14/2024
    if ($t -match '^\d{1,2}-\d{1,2}-\d{2,4}$') { return $null }  # 03-14-2024
    if ($t -match '^\(?\d{3}\)?[-. ]\d{3}[-. ]\d{4}$') { return $null }  # phone
    if ($t -match '\d\s*-\s*\d') { return $null }                # a range

    $neg = $false
    if ($t.StartsWith('(') -and $t.EndsWith(')')) { $neg = $true }

    $m = $script:InvRxMoneyUS.Match($t)
    $euro = $false
    if (-not $m.Success) {
        $m = $script:InvRxMoneyEU.Match($t)
        if (-not $m.Success) { return $null }
        $euro = $true
    }

    $n = $m.Groups['n'].Value
    if ($euro) { $n = $n.Replace('.', '').Replace(',', '.') }
    else       { $n = $n.Replace(',', '') }

    # A long run of digits with no decimal point is an identifier, not an amount.
    if ($n.IndexOf('.') -lt 0 -and $n.Length -gt 9) { return $null }

    $d = 0.0
    if (-not [double]::TryParse($n, [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $null }

    if (-not $euro -and $m.Groups['sfx'].Success) {
        $sfx = $m.Groups['sfx'].Value
        if ($sfx -eq 'CR' -or $sfx -eq '-') { $neg = $true }
    }
    if ($neg) { $d = -$d }

    if ([Math]::Abs($d) -ge 1000000000) { return $null }          # not a real invoice amount
    return $d
}

<#
 Parses a percentage ("7.0%", "7%", "(7.25 %)"). Returns [double] as a
 PERCENT (7.0 means 7%), or $null. A bare number is refused - without the
 sign there is no way to know whether 7 means 7% or $7.
#>
function Convert-InvPercent {
    [CmdletBinding()]
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $m = $script:InvRxPercent.Match($Text.Trim())
    if (-not $m.Success) { return $null }
    $d = 0.0
    if (-not [double]::TryParse($m.Groups['n'].Value, [System.Globalization.NumberStyles]::Float,
            [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $null }
    if ($d -lt 0 -or $d -gt 40) { return $null }                  # no real sales tax rate is above 40%
    return $d
}

<#
 Parses a date. Returns a [datetime] or $null.

 Handles the forms invoices actually use, including longhand months. Refuses
 anything outside 1990..(today + 1 year) - a "date" outside that window is a
 misread, not a date.

 Result carries an Ambiguous note when the text could be read either as
 month/day or day/month (e.g. 04/03/2024). Callers surface that rather than
 silently picking one.
#>
function Convert-InvDate {
    [CmdletBinding()]
    param([string]$Text, [switch]$DayFirst)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = $Text.Trim().TrimEnd('.', ',', ';', ':')
    $t = [regex]::Replace($t, '\s+', ' ')

    $lo = [datetime]'1990-01-01'
    $hi = (Get-Date).Date.AddYears(1)

    $mk = {
        param($y, $m, $d, $amb)
        if ($y -lt 100) { if ($y -ge 70) { $y += 1900 } else { $y += 2000 } }
        if ($m -lt 1 -or $m -gt 12 -or $d -lt 1 -or $d -gt 31) { return $null }
        $dt = $null
        try { $dt = Get-Date -Year $y -Month $m -Day $d -Hour 0 -Minute 0 -Second 0 -Millisecond 0 }
        catch { return $null }
        if ($dt -lt $lo -or $dt -gt $hi) { return $null }
        [pscustomobject]@{ Date = $dt; Ambiguous = [bool]$amb }
    }

    # 2024-03-14 / 2024/03/14
    $m = [regex]::Match($t, '^(?<y>\d{4})[-/.](?<m>\d{1,2})[-/.](?<d>\d{1,2})$')
    if ($m.Success) { return (& $mk ([int]$m.Groups['y'].Value) ([int]$m.Groups['m'].Value) ([int]$m.Groups['d'].Value) $false) }

    # 03/14/2024, 3-14-24, 03.14.2024
    $m = [regex]::Match($t, '^(?<a>\d{1,2})[-/.](?<b>\d{1,2})[-/.](?<y>\d{2,4})$')
    if ($m.Success) {
        $a = [int]$m.Groups['a'].Value
        $b = [int]$m.Groups['b'].Value
        $y = [int]$m.Groups['y'].Value
        if ($a -gt 12 -and $b -le 12) { return (& $mk $y $b $a $false) }      # must be d/m
        if ($b -gt 12 -and $a -le 12) { return (& $mk $y $a $b $false) }      # must be m/d
        $amb = ($a -le 12 -and $b -le 12)
        if ($DayFirst) { return (& $mk $y $b $a $amb) }
        return (& $mk $y $a $b $amb)
    }

    # 14-Mar-2024 / 14 March 2024
    $m = [regex]::Match($t, '^(?<d>\d{1,2})[- ](?<mon>[A-Za-z]{3,9})\.?[- ,]+(?<y>\d{2,4})$')
    if ($m.Success) {
        $mon = Get-InvMonthNumber $m.Groups['mon'].Value
        if ($mon) { return (& $mk ([int]$m.Groups['y'].Value) $mon ([int]$m.Groups['d'].Value) $false) }
    }

    # March 14, 2024 / Mar 14 2024
    $m = [regex]::Match($t, '^(?<mon>[A-Za-z]{3,9})\.?[- ]+(?<d>\d{1,2})(?:st|nd|rd|th)?[, ]+(?<y>\d{2,4})$')
    if ($m.Success) {
        $mon = Get-InvMonthNumber $m.Groups['mon'].Value
        if ($mon) { return (& $mk ([int]$m.Groups['y'].Value) $mon ([int]$m.Groups['d'].Value) $false) }
    }

    # 20240314
    $m = [regex]::Match($t, '^(?<y>(?:19|20)\d{2})(?<m>\d{2})(?<d>\d{2})$')
    if ($m.Success) { return (& $mk ([int]$m.Groups['y'].Value) ([int]$m.Groups['m'].Value) ([int]$m.Groups['d'].Value) $false) }

    return $null
}

$script:InvMonthTable = @{
    'jan'=1;'january'=1;'feb'=2;'february'=2;'mar'=3;'march'=3;'apr'=4;'april'=4
    'may'=5;'jun'=6;'june'=6;'jul'=7;'july'=7;'aug'=8;'august'=8
    'sep'=9;'sept'=9;'september'=9;'oct'=10;'october'=10;'nov'=11;'november'=11
    'dec'=12;'december'=12
}

function Get-InvMonthNumber {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $k = $Name.ToLowerInvariant().TrimEnd('.')
    if ($script:InvMonthTable.ContainsKey($k)) { return $script:InvMonthTable[$k] }
    return $null
}

<#
 Is this text a plausible invoice NUMBER?

 Strict on purpose: an invoice number must contain a digit, must not be a date
 or a money amount, must be 2..32 characters, and must not be a bare year or a
 lone small integer (those are page numbers and line counts).
#>
function Test-InvNumberLike {
    [CmdletBinding()]
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $t = $Text.Trim()

    if ($t.Length -lt 2 -or $t.Length -gt 32) { return $false }
    if ($t -notmatch '\d') { return $false }
    if ($t -match '^[A-Za-z]+$') { return $false }
    if ((Convert-InvDate $t)) { return $false }
    if ($t -match '^\$') { return $false }
    if ($t -match '^\d{1,3}(,\d{3})+(\.\d{2})?$') { return $false }     # 1,234.56
    if ($t -match '^\d+\.\d{2}$') { return $false }                     # 42.00
    if ($t -match '^(19|20)\d{2}$') { return $false }                   # a year
    if ($t -match '^\d{1,2}$') { return $false }                        # page number
    if ($t -match '^\d{3}[-. ]?\d{3}[-. ]?\d{4}$') { return $false }    # phone
    if ($t -match '%') { return $false }

    # "Invoice Total 684.80" reaches here as a candidate for the invoice NUMBER.
    # A FORMATTED amount inside the text disqualifies it - but only a formatted
    # one. A bare run of digits like 884213 is a perfectly ordinary invoice
    # number, so require a decimal, a thousands separator or a currency symbol
    # before rejecting.
    foreach ($tok in @(Get-InvMoneyTokens $t)) {
        if ($tok.Text -match '[.,$]') { return $false }
    }

    return $true
}

<#
 Normalises an invoice number for use as a key: upper case, and the noise
 characters that vary between printings removed. "inv# 4471" and "INV-4471"
 become the same key without changing what is reported.
#>
function ConvertTo-InvNumberKey {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return ([regex]::Replace($Text.ToUpperInvariant(), '[^A-Z0-9]', ''))
}

<#
 Finds every money amount in a piece of text, left to right.
 Used where a label and its value share a line with other numbers, e.g.
 "NE Sales Tax 7.0%: 62.16" -> the 62.16, not the 7.0.
#>
function Get-InvMoneyTokens {
    [CmdletBinding()]
    param([string]$Text)

    $out = [System.Collections.ArrayList]::new()
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }

    foreach ($m in [regex]::Matches($Text, '\(?\$?\s*\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?\)?|\(?\$?\s*\d+\.\d{1,2}\)?|\(?\$?\s*\d+\)?')) {
        $piece = $m.Value.Trim()

        # What follows decides whether this is money at all. "7.0%" must not
        # yield 7, and "03/14" must not yield 3 - so look past any digits and
        # separators still attached to the match before judging.
        $after = $Text.Substring([Math]::Min($Text.Length, $m.Index + $m.Length))
        if ($after -match '^[\d.,]*\s*%') { continue }          # a percentage
        if ($after -match '^\s*[/]') { continue }                # a date
        if ($after -match '^\d') { continue }                    # a truncated number
        if ($after -match '^-\d') { continue }                   # part of an identifier: 2024-0912

        # What precedes it matters too: the 2024 inside "PE-2024-4471" is not money.
        # The pattern may absorb leading spaces, so measure from the first real
        # character of the number, not from the start of the match.
        $lead = 0
        while ($lead -lt $m.Value.Length -and [char]::IsWhiteSpace($m.Value[$lead])) { $lead++ }
        $numAt = $m.Index + $lead
        if ($numAt -gt 0) {
            $before = $Text.Substring(0, $numAt)
            if ($before -match '[A-Za-z0-9]\s*[-/]$') { continue }
            if ($before -match '[A-Za-z0-9]$') { continue }
        }

        $v = Convert-InvMoney $piece
        if ($null -ne $v) {
            [void]$out.Add([pscustomobject]@{ Value = $v; Text = $piece; Index = $numAt })
        }
    }
    return @($out)
}
