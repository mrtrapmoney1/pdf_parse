<#
================================================================================
 InvRules.ps1
 The strict rules. This is what stands between "it parsed something" and
 "the number in the cell is right".

 TWO KINDS OF RULE
   FORMAT rules police one field on its own - a state must be a real state, a
   date must be inside a sane window, an amount must not be absurd. A field
   that fails a format rule is BLANKED and the reason recorded. A blank cell
   with an explanation is always better than a plausible wrong number.

   RECONCILIATION rules police the fields against each other - subtotal plus
   freight plus tax must equal the total, and the tax rate must agree with the
   tax divided by the taxable amount. These can also REPAIR: when exactly one
   value in an equation is missing, it is derived from the others and marked as
   derived. When they simply disagree, nothing is silently changed; the row is
   flagged for review with the arithmetic spelled out.

 Every rule writes a plain-English note. Nothing is failed silently.
================================================================================
#>

Set-StrictMode -Version 2.0

# How close two money values must be to count as agreeing. Invoices round to
# the cent, and rounding on a rate calculation can legitimately be a cent or
# two out on a large amount.
$script:InvMoneyTolerance = 0.02
$script:InvRateTolerance  = 0.15      # percentage points

function Add-InvIssue {
    param($Record, [string]$Text, [string]$Field = '')
    $msg = $(if ($Field) { $Field + ': ' + $Text } else { $Text })
    [void]$Record.Issues.Add($msg)
}

function Set-InvBlank {
    param($Record, [string]$Field, [string]$Why)
    if (-not $Record.Fields.ContainsKey($Field)) { return }
    $Record.Fields[$Field] = New-InvFinding -Value $null -Confidence 0 -Note $Why
    Add-InvIssue -Record $Record -Field $Field -Text ('cleared - ' + $Why)
}

function Get-InvValue {
    param($Record, [string]$Field)
    if (-not $Record.Fields.ContainsKey($Field)) { return $null }
    return $Record.Fields[$Field].Value
}

function Set-InvDerived {
    param($Record, [string]$Field, $Value, [int]$Confidence, [string]$How)
    $Record.Fields[$Field] = New-InvFinding -Value $Value -Confidence $Confidence `
                             -Source 'derived' -Note $How
    Add-InvIssue -Record $Record -Field $Field -Text ('derived - ' + $How)
}

<#
 Applies every rule to a record. Returns the record.
#>
function Test-InvRecord {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Record, $Identity = $null)

    $F = $Record.Fields
    if ($null -eq $F -or $F.Count -eq 0) { return $Record }

    # ---------------------------------------------------------------- format

    # money fields: sane sign and magnitude
    foreach ($fld in @('Subtotal','TaxAmt','Taxable','Freight','InvAmt')) {
        $v = Get-InvValue $Record $fld
        if ($null -eq $v) { continue }
        if ($v -lt 0) {
            # a credit memo is legitimate, but only when the TOTAL is negative too
            $tot = Get-InvValue $Record 'InvAmt'
            if ($fld -ne 'InvAmt' -and ($null -eq $tot -or $tot -ge 0)) {
                Set-InvBlank $Record $fld 'a negative amount here does not match a positive invoice total'
                continue
            }
        }
        if ([Math]::Abs($v) -gt 100000000) {
            Set-InvBlank $Record $fld 'the amount read is too large to be real'
        }
    }

    # tax rate must be a plausible sales-tax rate
    $rate = Get-InvValue $Record 'TaxRate'
    if ($null -ne $rate -and ($rate -lt 0 -or $rate -gt 25)) {
        Set-InvBlank $Record 'TaxRate' ('a rate of ' + $rate + '% is not a sales-tax rate')
    }

    # states must be real states
    foreach ($fld in @('State','ShipToState','BillToState')) {
        $v = Get-InvValue $Record $fld
        if ([string]::IsNullOrWhiteSpace([string]$v)) { continue }
        if (-not $InvStateSet.ContainsKey(([string]$v).ToUpperInvariant())) {
            Set-InvBlank $Record $fld ('"' + $v + '" is not a US state code')
        }
    }

    # dates must be inside a sane window
    $invDate = Get-InvValue $Record 'InvDate'
    if ($null -ne $invDate) {
        if ($invDate -gt (Get-Date).Date.AddDays(30)) {
            Set-InvBlank $Record 'InvDate' 'the invoice date is in the future'
            $invDate = $null
        }
        elseif ($invDate -lt [datetime]'1995-01-01') {
            Set-InvBlank $Record 'InvDate' 'the invoice date is implausibly old'
            $invDate = $null
        }
    }
    $dueDate = Get-InvValue $Record 'DueDate'
    if ($null -ne $dueDate -and $null -ne $invDate -and $dueDate -lt $invDate) {
        Add-InvIssue -Record $Record -Field 'DueDate' -Text 'the due date is before the invoice date'
        $Record.Fields['DueDate'].Confidence = [Math]::Min($Record.Fields['DueDate'].Confidence, 40)
    }

    # the invoice number must still look like one
    $num = Get-InvValue $Record 'InvNum'
    if (-not [string]::IsNullOrWhiteSpace([string]$num)) {
        if (-not (Test-InvNumberLike ([string]$num))) {
            Set-InvBlank $Record 'InvNum' ('"' + $num + '" is not a usable invoice number')
        }
    }

    # THE identity rule: we can never be the vendor
    $vend = Get-InvValue $Record 'Vendor'
    if ($null -ne $Identity -and -not [string]::IsNullOrWhiteSpace([string]$vend)) {
        if ((Test-InvIsSelf -Identity $Identity -Text ([string]$vend)) -ge 85) {
            Set-InvBlank $Record 'Vendor' 'this is our own company, not the vendor'
            foreach ($k in @('VendorAddress','City','State','Zip')) {
                Set-InvBlank $Record $k 'cleared with the vendor name'
            }
        }
    }

    # -------------------------------------------------------- reconciliation

    $sub  = Get-InvValue $Record 'Subtotal'
    $tax  = Get-InvValue $Record 'TaxAmt'
    $frt  = Get-InvValue $Record 'Freight'
    $disc = Get-InvValue $Record 'Discount'
    $tot  = Get-InvValue $Record 'InvAmt'
    $txbl = Get-InvValue $Record 'Taxable'

    $frtV  = $(if ($null -ne $frt)  { $frt }  else { 0.0 })
    $discV = $(if ($null -ne $disc) { [Math]::Abs($disc) } else { 0.0 })

    # subtotal + freight - discount + tax = total
    if ($null -ne $sub -and $null -ne $tax -and $null -ne $tot) {
        $calc = $sub + $frtV - $discV + $tax
        $diff = [Math]::Round($calc - $tot, 2)
        if ([Math]::Abs($diff) -gt $script:InvMoneyTolerance) {
            Add-InvIssue -Record $Record -Text (
                'the amounts do not add up: subtotal {0:N2} + freight {1:N2} - discount {2:N2} + tax {3:N2} = {4:N2}, but the total reads {5:N2} (out by {6:N2})' `
                -f $sub, $frtV, $discV, $tax, $calc, $tot, $diff)
            foreach ($k in @('Subtotal','TaxAmt','InvAmt')) {
                $Record.Fields[$k].Confidence = [Math]::Min($Record.Fields[$k].Confidence, 45)
            }
        }
        else {
            # the arithmetic proves all three; this is the strongest evidence available
            foreach ($k in @('Subtotal','TaxAmt','InvAmt')) {
                $Record.Fields[$k].Confidence = [Math]::Min(99, $Record.Fields[$k].Confidence + 10)
            }
            Add-InvIssue -Record $Record -Text 'checked: subtotal + freight + tax equals the invoice total'
        }
    }
    elseif ($null -ne $tot -and $null -ne $tax -and $null -eq $sub) {
        Set-InvDerived $Record 'Subtotal' ([Math]::Round($tot - $tax - $frtV + $discV, 2)) 62 `
            'total minus tax and freight - the invoice showed no subtotal'
        $sub = Get-InvValue $Record 'Subtotal'
    }
    elseif ($null -ne $tot -and $null -ne $sub -and $null -eq $tax) {
        $d = [Math]::Round($tot - $sub - $frtV + $discV, 2)
        if ($d -ge 0 -and $d -le ($sub * 0.25)) {
            Set-InvDerived $Record 'TaxAmt' $d 58 'total minus subtotal and freight - the invoice showed no tax line'
            $tax = $d
        }
    }

    # taxable amount, tax and rate must agree
    if ($null -eq $txbl) {
        if ($null -ne $tax -and $null -ne $rate -and $rate -gt 0) {
            $t = [Math]::Round(($tax / ($rate / 100.0)), 2)
            if ($t -gt 0 -and ($null -eq $tot -or $t -le ($tot * 1.05))) {
                Set-InvDerived $Record 'Taxable' $t 60 ('tax {0:N2} divided by the stated rate {1}%' -f $tax, $rate)
                $txbl = $t
            }
        }
        elseif ($null -ne $sub) {
            # the common case: everything on the invoice was taxable
            Set-InvDerived $Record 'Taxable' $sub 45 'assumed equal to the subtotal - the invoice did not state a taxable amount'
            $txbl = $sub
        }
    }

    if ($null -ne $txbl -and $null -ne $tax -and $txbl -gt 0) {
        $calcRate = [Math]::Round(($tax / $txbl) * 100.0, 4)
        if ($null -eq $rate) {
            if ($calcRate -ge 0 -and $calcRate -le 25) {
                Set-InvDerived $Record 'TaxRate' ([Math]::Round($calcRate, 4)) 66 `
                    ('tax {0:N2} divided by taxable {1:N2}' -f $tax, $txbl)
            }
        }
        elseif ([Math]::Abs($calcRate - $rate) -gt $script:InvRateTolerance) {
            Add-InvIssue -Record $Record -Text (
                'the tax rate does not agree with the amounts: {0:N2} on {1:N2} is {2:N3}%, but the invoice states {3}%' `
                -f $tax, $txbl, $calcRate, $rate)
            $Record.Fields['TaxRate'].Confidence = [Math]::Min($Record.Fields['TaxRate'].Confidence, 45)
        }
        else {
            $Record.Fields['TaxRate'].Confidence = [Math]::Min(99, $Record.Fields['TaxRate'].Confidence + 10)
        }
    }

    if ($null -ne $txbl -and $null -ne $tot -and $txbl -gt ($tot + 0.01)) {
        Add-InvIssue -Record $Record -Field 'Taxable' -Text 'the taxable amount is larger than the invoice total'
        $Record.Fields['Taxable'].Confidence = [Math]::Min($Record.Fields['Taxable'].Confidence, 40)
    }

    # ------------------------------------------------------------- shipping

    # Ship-to block OR a freight charge is evidence that goods moved. Neither
    # is required though: plenty of invoices ship goods with no ship-to block
    # and no separate shipping line, so N means "no evidence on the page", not
    # "nothing shipped". The Notes column says which it is.
    $hasShipBlock = -not [string]::IsNullOrWhiteSpace([string](Get-InvValue $Record 'ShipToName')) -or
                    -not [string]::IsNullOrWhiteSpace([string](Get-InvValue $Record 'ShipToCity'))
    $hasFreight = ($null -ne $frt -and $frt -gt 0)

    if ($hasShipBlock -and $hasFreight) {
        $F['Shipped'] = New-InvFinding -Value 'Y' -Confidence 95 -Source 'ship-to block and a freight charge'
    }
    elseif ($hasShipBlock) {
        $F['Shipped'] = New-InvFinding -Value 'Y' -Confidence 88 -Source 'ship-to block'
    }
    elseif ($hasFreight) {
        $F['Shipped'] = New-InvFinding -Value 'Y' -Confidence 72 -Source 'a freight charge, but no ship-to block'
        Add-InvIssue -Record $Record -Text 'freight was charged but the invoice shows no ship-to address, so the delivery jurisdiction is unknown'
    }
    else {
        $F['Shipped'] = New-InvFinding -Value 'N' -Confidence 60 -Source 'no ship-to block and no freight charge'
    }

    # ------------------------------------------------------------- required

    $required = @('Vendor','InvNum','InvDate','InvAmt')
    foreach ($r in $required) {
        if ([string]::IsNullOrWhiteSpace([string](Get-InvValue $Record $r))) {
            Add-InvIssue -Record $Record -Text ('no ' + $r + ' could be found')
        }
    }

    # -------------------------------------------- row confidence and review

    $key = @('Vendor','InvNum','InvDate','InvAmt','TaxAmt')
    $low = 100
    foreach ($k in $key) {
        if (-not $F.ContainsKey($k)) { continue }
        $c = $F[$k].Confidence
        if ($null -eq (Get-InvValue $Record $k)) { $c = 0 }
        if ($c -lt $low) { $low = $c }
    }
    $Record.RowConfidence = $low
    $Record.NeedsReview = ($low -lt 75)

    # VenInvoice: the workpaper key, vendor and number together
    $vk = [string](Get-InvValue $Record 'Vendor')
    $nk = [string](Get-InvValue $Record 'InvNum')
    $F['VenInvoice'] = New-InvFinding -Value (($vk + ' ' + $nk).Trim()) `
                        -Confidence $low -Source 'vendor + invoice number'

    return $Record
}
