Start-InvSuite 'Edge cases - the files that will actually turn up in a folder'

$id = Get-InvIdentity -Path (Join-Path $script:InvRoot 'config/MyCompany.json')

# --- a scan: pages, but no text layer ---------------------------------------
$rec = Get-InvRecord -Path (Join-Path $script:InvRoot 'samples/scan_no_text.pdf') -Identity $id
Assert-InvEqual 'a scan reports Provided = N' 'N' $rec.Provided
Assert-InvTrue  'a scan is flagged for review' $rec.NeedsReview
Assert-InvTrue  'a scan says it needs OCR' ((@($rec.Issues) -join ' ') -match 'OCR')

# --- a file that is not a PDF at all ----------------------------------------
$notPdf = Join-Path ([System.IO.Path]::GetTempPath()) 'inv_not_really.pdf'
'this is not a pdf' | Set-Content -LiteralPath $notPdf -Encoding UTF8
$rec = Get-InvRecord -Path $notPdf -Identity $id
Assert-InvEqual 'a non-PDF reports Provided = N' 'N' $rec.Provided
Assert-InvTrue  'and says why' ((@($rec.Issues) -join ' ') -match 'damaged|not a PDF')
Remove-Item -LiteralPath $notPdf -Force -ErrorAction SilentlyContinue

# --- a missing file ----------------------------------------------------------
$rec = Get-InvRecord -Path (Join-Path $script:InvRoot 'samples/does_not_exist.pdf') -Identity $id
Assert-InvEqual 'a missing file reports Provided = N' 'N' $rec.Provided

# --- OUR OWN letterhead: a credit memo we issued, misfiled into the folder ---
# The identity guard must veto us, and then must NOT fall through to reporting
# a field label as the vendor.
$rec = Get-InvRecord -Path (Join-Path $script:InvRoot 'samples/self_letterhead.pdf') -Identity $id
$v = [string]$rec.Fields['Vendor'].Value
Assert-InvTrue 'our own name is never the vendor' `
    ((Test-InvIsSelf -Identity $id -Text $v) -lt 85)
# The rule is REFUSE, not "pick the next best line". Anything non-empty here -
# a field label, a city, a document title - is a wrong answer.
Assert-InvNull 'nothing at all is reported as the vendor' $v
Assert-InvTrue 'and the reason says it is our own letterhead' `
    ($rec.Fields['Vendor'].Note -match 'our own company')
Assert-InvTrue 'the row is flagged' $rec.NeedsReview

# a city line must never pass as a company name, anywhere
Assert-InvTrue 'a city/state/zip line is not a company' (-not (Test-InvCompanyLine 'Omaha, NE 68102'))
Assert-InvTrue 'a street line is not a company'         (-not (Test-InvCompanyLine '123 Main Street'))
Assert-InvTrue 'a field label is not a company'         (-not (Test-InvCompanyLine 'Invoice #:'))
Assert-InvTrue 'a real company still is one'            (Test-InvCompanyLine 'Midwest Supply Co.')
Assert-InvTrue 'and one whose name starts with a label word' (Test-InvCompanyLine 'Total Wine & More')

# --- amounts that do not reconcile ------------------------------------------
$rec = Get-InvRecord -Path (Join-Path $script:InvRoot 'samples/bad_math.pdf') -Identity $id
Assert-InvEqual 'the vendor is still read'  'TRI-STATE FASTENERS' $rec.Fields['Vendor'].Value
Assert-InvEqual 'the stated total is kept as printed' 999.99 $rec.Fields['InvAmt'].Value
Assert-InvEqual 'the stated subtotal is kept as printed' 500.00 $rec.Fields['Subtotal'].Value
Assert-InvTrue  'the mismatch is spelled out' `
    ((@($rec.Issues) -join ' ') -match 'do not add up')
Assert-InvTrue  'and the row is flagged' $rec.NeedsReview
Assert-InvTrue  'confidence is knocked down' ($rec.RowConfidence -le 50)
