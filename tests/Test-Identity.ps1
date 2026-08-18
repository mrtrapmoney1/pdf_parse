Start-InvSuite 'Identity - we are never the vendor'

# name keys survive punctuation and legal-suffix differences
Assert-InvEqual 'llc dropped'        'acmeholdings' (ConvertTo-InvNameKey 'Acme Holdings LLC')
Assert-InvEqual 'dotted llc dropped' 'acmeholdings' (ConvertTo-InvNameKey 'ACME HOLDINGS, L.L.C.')
Assert-InvEqual 'inc dropped'        'acmeholdings' (ConvertTo-InvNameKey 'Acme Holdings, Inc.')
Assert-InvTrue  'holdings is part of the name, not a suffix' `
                ((ConvertTo-InvNameKey 'Acme Holdings LLC') -ne 'acme')

$id = Get-InvIdentity -Path (Join-Path $script:InvRoot 'config/MyCompany.json')
Assert-InvTrue 'the identity file loads' ($null -ne $id)

Assert-InvTrue 'exact name is us'      ((Test-InvIsSelf -Identity $id -Text 'Acme Holdings LLC') -ge 95)
Assert-InvTrue 'alias is us'           ((Test-InvIsSelf -Identity $id -Text 'ACME HOLDINGS, L.L.C.') -ge 95)
Assert-InvTrue 'name inside a line is us' ((Test-InvIsSelf -Identity $id -Text 'Sold To: Acme Holdings LLC') -ge 90)
Assert-InvEqual 'a real vendor is not us' 0 (Test-InvIsSelf -Identity $id -Text 'Midwest Supply Co.')
Assert-InvEqual 'a similar-but-different name is not us' 0 (Test-InvIsSelf -Identity $id -Text 'Acme Plumbing Supply')

# bill-to and ship-to printed side by side must not interleave
$r = Get-InvPdfWords -Path (Join-Path $script:InvRoot 'samples/inv_classic.pdf')
$L = New-InvLayout -Words $r.Words -PageInfo $r.PageInfo
$blocks = @(Get-InvPartyBlocks -Layout $L -Identity $id)

$bill = @($blocks | Where-Object { $_.Kind -eq 'BillTo' })
$ship = @($blocks | Where-Object { $_.Kind -eq 'ShipTo' })
Assert-InvEqual 'one bill-to block' 1 $bill.Count
Assert-InvEqual 'one ship-to block' 1 $ship.Count
Assert-InvEqual 'bill-to is clean'  'Acme Holdings LLC' $bill[0].Lines[0]
Assert-InvEqual 'ship-to is clean'  'Acme Warehouse #4' $ship[0].Lines[0]
Assert-InvTrue  'bill-to is us'     ($bill[0].IsSelf -ge 90)
Assert-InvTrue  'ship-to is us'     ($ship[0].IsSelf -ge 90)

# the vendor is the letterhead, never the customer
$v = Get-InvVendor -Layout $L -Identity $id -Blocks $blocks
Assert-InvEqual 'vendor is the letterhead' 'MIDWEST SUPPLY CO.' $v.Name.Value
Assert-InvEqual 'vendor city'  'Lincoln' $v.City.Value
Assert-InvEqual 'vendor state' 'NE'      $v.State.Value

# addresses in every shape they turn up in
$cases = @(
    @{ Lines = @('X','4820 South 72nd Street','Lincoln, NE 68516');  City='Lincoln';        State='NE'; Zip='68516' }
    @{ Lines = @('X','1201 W Adams | Omaha, NE 68132');              City='Omaha';          State='NE'; Zip='68132' }
    @{ Lines = @('X','77 Riverfront Drive, Council Bluffs, IA 51501');City='Council Bluffs'; State='IA'; Zip='51501' }
    @{ Lines = @('X','PO Box 4410, Kearney, NE 68848');              City='Kearney';        State='NE'; Zip='68848' }
    @{ Lines = @('X','Suite 200 Omaha NE 68102');                    City='Omaha';          State='NE'; Zip='68102' }
    @{ Lines = @('X','900 Elm St','Grand Island, Nebraska 68801');    City='Grand Island';   State='NE'; Zip='68801' }
)
foreach ($c in $cases) {
    $p = Get-InvAddressParts -Lines $c.Lines
    Assert-InvEqual ('city from <' + $c.Lines[-1] + '>')  $c.City  $p.City
    Assert-InvEqual ('state from <' + $c.Lines[-1] + '>') $c.State $p.State
}
