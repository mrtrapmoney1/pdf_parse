Start-InvSuite 'Excel output'

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('invtest_' + [Guid]::NewGuid().ToString('N').Substring(0,8) + '.xlsx')

$cols = Resolve-InvColumns -Keys @('Vendor','InvDate','TaxRate','InvAmt')
$rows = @(
    @{ VenInvoice=@{V='ACME 1'}; Source=@{V='a.pdf'}; Provided=@{V='Y'}
       Vendor=@{V='Midwest Supply'}; InvDate=@{V=[datetime]'2024-03-14'}
       TaxRate=@{V=7.0}; InvAmt=@{V=950.16} }
    @{ VenInvoice=@{V='BLUFF 2'}; Source=@{V='b.pdf'}; Provided=@{V='Y'}
       Vendor=@{V='Bluff City';Low=$true}; InvDate=@{V=[datetime]'2024-05-21'}
       TaxRate=@{V=7.0}; InvAmt=@{V=684.80} }
)

$p = Export-InvWorkbook -Path $tmp -Columns $cols -Rows $rows `
        -LogLines @([pscustomobject]@{File='a.pdf';Message='ok'})

Assert-InvTrue 'the workbook was created' (Test-Path -LiteralPath $p)
Assert-InvTrue 'it is not empty' ((Get-Item -LiteralPath $p).Length -gt 1000)

# it must be a readable zip holding the parts Excel needs
Initialize-InvZip
$zip = [System.IO.Compression.ZipFile]::OpenRead($p)
try {
    $names = @($zip.Entries | ForEach-Object { $_.FullName })
    foreach ($need in @('[Content_Types].xml','_rels/.rels','xl/workbook.xml',
                        'xl/_rels/workbook.xml.rels','xl/styles.xml',
                        'xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml')) {
        Assert-InvTrue ('part present: ' + $need) ($names -contains $need)
    }

    $entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
    $sr = [System.IO.StreamReader]::new($entry.Open())
    $sheet = $sr.ReadToEnd()
    $sr.Dispose()

    # every part must be well-formed XML or Excel refuses the whole file
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($sheet)
    Assert-InvTrue 'the sheet is well-formed xml' ($null -ne $doc.DocumentElement)

    Assert-InvTrue 'the header row is frozen'  ($sheet -match 'state="frozen"')
    Assert-InvTrue 'a filter is applied'       ($sheet -match '<autoFilter')
    Assert-InvTrue 'dates are written as numbers, not text' ($sheet -match '<v>45365</v>')
    Assert-InvTrue 'percentages are stored as fractions'    ($sheet -match '<v>0\.07</v>')
    Assert-InvTrue 'money is a number'                      ($sheet -match '<v>950\.16</v>')
    Assert-InvTrue 'low-confidence cells use the amber style' ($sheet -match 's="7"')
}
finally { $zip.Dispose() }

# XML-hostile text must not be able to corrupt the file
$rows2 = @(@{ VenInvoice=@{V='x'}; Source=@{V='<&>"'}; Provided=@{V='Y'}
              Vendor=@{V='Smith & Sons <Ltd>'}; InvAmt=@{V=1.0} })
$p2 = Export-InvWorkbook -Path ([System.IO.Path]::ChangeExtension($tmp, '.esc.xlsx')) -Columns $cols -Rows $rows2
$zip2 = [System.IO.Compression.ZipFile]::OpenRead($p2)
try {
    $e2 = $zip2.GetEntry('xl/worksheets/sheet1.xml')
    $sr2 = [System.IO.StreamReader]::new($e2.Open())
    $x2 = $sr2.ReadToEnd(); $sr2.Dispose()
    $d2 = New-Object System.Xml.XmlDocument
    $d2.LoadXml($x2)
    Assert-InvTrue 'ampersands and angle brackets are escaped' ($x2 -match 'Smith &amp; Sons &lt;Ltd&gt;')
}
finally { $zip2.Dispose() }

Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $p2 -Force -ErrorAction SilentlyContinue
