Start-InvSuite 'End to end - the sample invoices'

$id = Get-InvIdentity -Path (Join-Path $script:InvRoot 'config/MyCompany.json')

# Each of these was checked by hand against the PDF it comes from. They are the
# regression baseline: if a change breaks one of these, it broke extraction.
$expected = @(
    @{ File='inv_classic.pdf';      Vendor='MIDWEST SUPPLY CO.';           City='Lincoln';        State='NE'
       InvNum='MS-88213';     InvDate='03/14/2024'; Subtotal=853.00;  TaxAmt=62.16; InvAmt=950.16
       Freight=35.00; TaxRate=7.0; ShipToCity='Grand Island'; ShipToState='NE'; Shipped='Y' }

    @{ File='inv_labels_above.pdf'; Vendor='Prairie Electric, Inc.';       City='Omaha';          State='NE'
       InvNum='PE-2024-4471'; InvDate='04/02/2024'; Taxable=1203.00; TaxAmt=84.45; InvAmt=1287.45
       TaxRate=7.0; Shipped='N' }

    @{ File='inv_shipto.pdf';       Vendor='BLUFF CITY HARDWARE';          City='Council Bluffs'; State='IA'
       InvNum='BC-5590';      InvDate='05/21/2024'; Subtotal=640.00;  TaxAmt=44.80; InvAmt=684.80
       TaxRate=7.0; ShipToCity='Lincoln'; ShipToState='NE'; Shipped='Y' }

    @{ File='inv_twopage.pdf';      Vendor='CORNHUSKER INDUSTRIAL SUPPLY'; City='Kearney';        State='NE'
       InvNum='2024-0912';    InvDate='06/30/2024'; Subtotal=516.00;  TaxAmt=28.38; InvAmt=544.38
       TaxRate=5.5; Shipped='N' }
)

foreach ($e in $expected) {
    $path = Join-Path $script:InvRoot ('samples/' + $e.File)
    $rec = Get-InvRecord -Path $path -Identity $id

    Assert-InvEqual ($e.File + ' : text was read') 'Y' $rec.Provided

    foreach ($k in $e.Keys) {
        if ($k -eq 'File') { continue }
        $actual = $rec.Fields[$k].Value
        if ($k -eq 'InvDate' -and $null -ne $actual) { $actual = ([datetime]$actual).ToString('MM/dd/yyyy') }
        Assert-InvEqual ($e.File + ' : ' + $k) $e[$k] $actual
    }

    # our own company must never appear in the vendor column
    Assert-InvTrue ($e.File + ' : vendor is not us') `
        ((Test-InvIsSelf -Identity $id -Text ([string]$rec.Fields['Vendor'].Value)) -lt 85)
}
