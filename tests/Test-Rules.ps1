Start-InvSuite 'Rules - format policing and reconciliation'

function New-InvFakeRecord {
    param([hashtable]$Values)
    $F = @{}
    foreach ($k in $Values.Keys) {
        $F[$k] = New-InvFinding -Value $Values[$k] -Confidence 80 -Source 'test'
    }
    foreach ($k in @('Vendor','InvNum','InvDate','InvAmt','TaxAmt','Taxable','Subtotal',
                     'Freight','Discount','TaxRate','State','ShipToState','ShipToName',
                     'ShipToCity','DueDate','City','Zip','VendorAddress')) {
        if (-not $F.ContainsKey($k)) { $F[$k] = New-InvFinding -Value $null -Confidence 0 }
    }
    [pscustomobject]@{
        SourcePath='x.pdf'; Source='x.pdf'; Provided='Y'; Pages=1; Scanned=$false
        Fields=$F; Issues=[System.Collections.ArrayList]::new()
        NeedsReview=$true; RowConfidence=0
    }
}

# --- the totals check --------------------------------------------------------
$r = New-InvFakeRecord @{ Subtotal=853.00; Freight=35.00; TaxAmt=62.16; InvAmt=950.16 }
$r = Test-InvRecord -Record $r
Assert-InvTrue 'consistent amounts are confirmed' `
    ((@($r.Issues) -join ' ') -match 'checked: subtotal')
Assert-InvTrue 'confirmed arithmetic raises confidence' ($r.Fields['InvAmt'].Confidence -ge 88)

$r = New-InvFakeRecord @{ Subtotal=853.00; TaxAmt=62.16; InvAmt=999.99 }
$r = Test-InvRecord -Record $r
Assert-InvTrue 'a mismatch is reported' ((@($r.Issues) -join ' ') -match 'do not add up')
Assert-InvTrue 'a mismatch lowers confidence' ($r.Fields['InvAmt'].Confidence -le 45)
Assert-InvEqual 'a mismatch does NOT silently change the total' 999.99 $r.Fields['InvAmt'].Value

# --- derivation --------------------------------------------------------------
$r = New-InvFakeRecord @{ TaxAmt=44.80; InvAmt=684.80 }
$r = Test-InvRecord -Record $r
Assert-InvEqual 'subtotal derived from total minus tax' 640.00 $r.Fields['Subtotal'].Value
Assert-InvTrue  'derivation is disclosed' ($r.Fields['Subtotal'].Note -match 'total minus tax')

$r = New-InvFakeRecord @{ TaxAmt=62.16; TaxRate=7.0 }
$r = Test-InvRecord -Record $r
Assert-InvEqual 'taxable derived from tax and rate' 888.00 $r.Fields['Taxable'].Value

$r = New-InvFakeRecord @{ TaxAmt=44.80; Taxable=640.00 }
$r = Test-InvRecord -Record $r
Assert-InvEqual 'rate derived from tax over taxable' 7.0 $r.Fields['TaxRate'].Value

# --- the rate cross-check ----------------------------------------------------
$r = New-InvFakeRecord @{ TaxAmt=44.80; Taxable=640.00; TaxRate=9.5 }
$r = Test-InvRecord -Record $r
Assert-InvTrue 'a disagreeing rate is reported' ((@($r.Issues) -join ' ') -match 'does not agree')

# --- format rules ------------------------------------------------------------
$r = New-InvFakeRecord @{ State='ZZ'; ShipToState='NE' }
$r = Test-InvRecord -Record $r
Assert-InvNull  'a bogus state is blanked' $r.Fields['State'].Value
Assert-InvEqual 'a real state survives'    'NE' $r.Fields['ShipToState'].Value

$r = New-InvFakeRecord @{ TaxRate=68.0 }
$r = Test-InvRecord -Record $r
Assert-InvNull 'an absurd tax rate is blanked' $r.Fields['TaxRate'].Value

$r = New-InvFakeRecord @{ InvDate=(Get-Date).Date.AddYears(1) }
$r = Test-InvRecord -Record $r
Assert-InvNull 'a future invoice date is blanked' $r.Fields['InvDate'].Value

$r = New-InvFakeRecord @{ InvNum='42.00' }
$r = Test-InvRecord -Record $r
Assert-InvNull 'an amount is not an invoice number' $r.Fields['InvNum'].Value

# --- the identity rule -------------------------------------------------------
$id = Get-InvIdentity -Path (Join-Path $script:InvRoot 'config/MyCompany.json')
$r = New-InvFakeRecord @{ Vendor='Acme Holdings LLC'; City='Omaha'; State='NE' }
$r = Test-InvRecord -Record $r -Identity $id
Assert-InvNull 'our own name is cleared from Vendor' $r.Fields['Vendor'].Value
Assert-InvNull 'and its city goes with it'           $r.Fields['City'].Value
Assert-InvTrue 'and the reason is recorded' ((@($r.Issues) -join ' ') -match 'our own company')

# --- shipping evidence -------------------------------------------------------
$r = New-InvFakeRecord @{ ShipToCity='Lincoln' }
$r = Test-InvRecord -Record $r
Assert-InvEqual 'a ship-to block means shipped' 'Y' $r.Fields['Shipped'].Value

$r = New-InvFakeRecord @{ Freight=35.00; Subtotal=100.00; InvAmt=135.00 }
$r = Test-InvRecord -Record $r
Assert-InvEqual 'freight alone still means shipped' 'Y' $r.Fields['Shipped'].Value
Assert-InvTrue  'but the unknown destination is flagged' `
    ((@($r.Issues) -join ' ') -match 'no ship-to address')

$r = New-InvFakeRecord @{ Subtotal=100.00; InvAmt=100.00 }
$r = Test-InvRecord -Record $r
Assert-InvEqual 'no evidence reads N' 'N' $r.Fields['Shipped'].Value
