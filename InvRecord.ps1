<#
================================================================================
 InvRecord.ps1
 One PDF in, one fully-populated record out.

 This is the only place that knows the whole shape of an invoice. It runs every
 extractor, hands the results to the rules engine, and returns a record whose
 every field carries a value, a confidence and a reason.
================================================================================
#>

Set-StrictMode -Version 2.0

function Get-InvRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        $Identity = $null,
        [switch] $NoRules
    )

    $file = Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
    $rec = [ordered]@{
        SourcePath = $Path
        Source     = $(if ($file) { $file.Name } else { Split-Path -Leaf $Path })
        Provided   = 'N'
        Pages      = 0
        Scanned    = $false
        Fields     = @{}
        Issues     = [System.Collections.ArrayList]::new()
        NeedsReview = $true
        RowConfidence = 0
    }

    $txt = Get-InvPdfWords -Path $Path
    $rec.Pages   = $txt.Pages
    $rec.Scanned = $txt.Scanned

    if (-not $txt.Ok) {
        [void]$rec.Issues.Add($(if ($txt.Note) { $txt.Note } else { 'No text could be read from this PDF' }))
        return [pscustomobject]$rec
    }

    $rec.Provided = 'Y'
    $L = New-InvLayout -Words $txt.Words -PageInfo $txt.PageInfo
    $blocks = @(Get-InvPartyBlocks -Layout $L -Identity $Identity)

    $F = @{}

    # --- who ---------------------------------------------------------------
    $vendor = Get-InvVendor -Layout $L -Identity $Identity -Blocks $blocks
    $F['Vendor']        = $vendor.Name
    $F['VendorAddress'] = $vendor.Address
    $F['City']          = $vendor.City
    $F['State']         = $vendor.State
    $F['Zip']           = $vendor.Zip

    # --- ship-to -----------------------------------------------------------
    # Captured whenever the block exists, freight or no freight: for use tax the
    # delivery address often decides the jurisdiction, and plenty of invoices
    # ship goods without a separate shipping line.
    $shipBlocks = @($blocks | Where-Object { $_.Kind -eq 'ShipTo' })
    if ($shipBlocks.Count -gt 0) {
        $sb = $shipBlocks[0]
        $sp = Get-InvAddressParts -Lines $sb.Lines
        $conf = [Math]::Max(40, [Math]::Min(96, $sb.Score))
        $F['ShipToName']    = New-InvFinding -Value $sp.Name  -Confidence $conf -Source ('block:ShipTo (' + $sb.Heading + ')')
        $F['ShipToAddress'] = New-InvFinding -Value (@($sb.Lines | Select-Object -Skip 1) -join ', ') -Confidence $conf -Source 'block:ShipTo'
        $F['ShipToCity']    = New-InvFinding -Value $sp.City  -Confidence $(if ($sp.City)  { [Math]::Min($conf, $sp.Confidence) } else { 0 }) -Source 'block:ShipTo'
        $F['ShipToState']   = New-InvFinding -Value $sp.State -Confidence $(if ($sp.State) { [Math]::Min($conf, $sp.Confidence) } else { 0 }) -Source 'block:ShipTo'
        $F['ShipToZip']     = New-InvFinding -Value $sp.Zip   -Confidence $(if ($sp.Zip)   { [Math]::Min($conf, $sp.Confidence) } else { 0 }) -Source 'block:ShipTo'
        $F['Shipped']       = New-InvFinding -Value 'Y' -Confidence $conf -Source 'ship-to block present'
    }
    else {
        foreach ($k in @('ShipToName','ShipToAddress','ShipToCity','ShipToState','ShipToZip')) {
            $F[$k] = New-InvFinding -Value $null -Confidence 0 -Note 'No ship-to block on this invoice'
        }
        $F['Shipped'] = New-InvFinding -Value '' -Confidence 0
    }

    $billBlocks = @($blocks | Where-Object { $_.Kind -eq 'BillTo' })
    if ($billBlocks.Count -gt 0) {
        $bp = Get-InvAddressParts -Lines $billBlocks[0].Lines
        $F['BillToName']  = New-InvFinding -Value $bp.Name  -Confidence $billBlocks[0].Score -Source 'block:BillTo'
        $F['BillToCity']  = New-InvFinding -Value $bp.City  -Confidence $bp.Confidence -Source 'block:BillTo'
        $F['BillToState'] = New-InvFinding -Value $bp.State -Confidence $bp.Confidence -Source 'block:BillTo'
    }
    else {
        foreach ($k in @('BillToName','BillToCity','BillToState')) {
            $F[$k] = New-InvFinding -Value $null -Confidence 0
        }
    }

    # --- identifiers -------------------------------------------------------
    $F['InvNum'] = Resolve-InvFinding (Get-InvTextField -Layout $L -Labels $InvLabels.InvNum `
                        -Validator { param($t) Test-InvNumberLike $t })
    $F['PONum']  = Resolve-InvFinding (Get-InvTextField -Layout $L -Labels $InvLabels.PONum `
                        -Validator { param($t) Test-InvNumberLike $t })
    $F['AcctNum'] = Resolve-InvFinding (Get-InvTextField -Layout $L -Labels $InvLabels.AcctNum `
                        -Validator { param($t) Test-InvNumberLike $t })
    $F['Terms']  = Resolve-InvFinding (Get-InvTextField -Layout $L -Labels $InvLabels.Terms `
                        -Validator { param($t) ($t.Length -ge 2 -and $t.Length -le 40 -and $t -match '[A-Za-z0-9]') } -MaxWords 4)

    # --- dates -------------------------------------------------------------
    $F['InvDate'] = Resolve-InvFinding (Get-InvDateField -Layout $L -Labels $InvLabels.InvDate)
    $F['DueDate'] = Resolve-InvFinding (Get-InvDateField -Layout $L -Labels $InvLabels.DueDate)

    # --- money -------------------------------------------------------------
    $F['Subtotal'] = Resolve-InvFinding (Get-InvMoneyField -Layout $L -Labels $InvLabels.Subtotal)
    $F['TaxAmt']   = Resolve-InvFinding (Get-InvMoneyField -Layout $L -Labels $InvLabels.TaxAmt)
    $F['Taxable']  = Resolve-InvFinding (Get-InvMoneyField -Layout $L -Labels $InvLabels.TaxableAmount)
    $F['Freight']  = Resolve-InvFinding (Get-InvMoneyField -Layout $L -Labels $InvLabels.Freight)
    $F['Discount'] = Resolve-InvFinding (Get-InvMoneyField -Layout $L -Labels $InvLabels.Discount)
    $F['InvAmt']   = Resolve-InvFinding (Get-InvMoneyField -Layout $L -Labels $InvLabels.InvAmt -PreferLastPage)

    # --- tax rate ----------------------------------------------------------
    # A stated rate is often printed on the tax LINE itself ("NE Sales Tax 7.0%")
    # rather than under a "Tax Rate" label, so look in both places.
    $rateFindings = @(Get-InvPercentField -Layout $L -Labels $InvLabels.TaxRate)
    foreach ($h in @(Find-InvLabel -Layout $L -Labels $InvLabels.TaxAmt)) {
        $m = [regex]::Match($h.Line.Text, '(\d{1,2}(?:\.\d{1,4})?)\s*%')
        if (-not $m.Success) { continue }
        $p = Convert-InvPercent $m.Value
        if ($null -eq $p) { continue }
        $rateFindings += (New-InvFinding -Value $p -Confidence ([Math]::Max(1, $h.Score - 8)) `
                          -Source 'percentage on the tax line' -Raw $m.Value -Page $h.Page)
        break
    }
    $F['TaxRate'] = Resolve-InvFinding $rateFindings

    # --- description -------------------------------------------------------
    $F['Description'] = Get-InvDescription -Layout $L -Identity $Identity

    $F['Currency'] = New-InvFinding -Value 'USD' -Confidence 40 -Source 'default'
    $F['LineCount'] = New-InvFinding -Value $null -Confidence 0
    $F['PageCount'] = New-InvFinding -Value $txt.Pages -Confidence 100 -Source 'pdf'

    $rec.Fields = $F
    $out = [pscustomobject]$rec

    if (-not $NoRules) { $out = Test-InvRecord -Record $out -Identity $Identity }
    return $out
}
