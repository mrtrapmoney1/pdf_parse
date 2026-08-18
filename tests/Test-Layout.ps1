Start-InvSuite 'Layout - keys, lines, label anchoring, column groups'

# the squash key collapses every spelling of the same label
Assert-InvEqual 'colon dropped'   'invoice#'  (ConvertTo-InvKey 'Invoice #:')
Assert-InvEqual 'no space'        'invoice#'  (ConvertTo-InvKey 'Invoice#')
Assert-InvEqual 'dots dropped'    'invoiceno' (ConvertTo-InvKey 'INVOICE NO.')
Assert-InvEqual 'dash dropped'    'billto'    (ConvertTo-InvKey 'Bill-To')
Assert-InvEqual 'space dropped'   'billto'    (ConvertTo-InvKey 'Bill To')
Assert-InvEqual 'subtotal forms'  'subtotal'  (ConvertTo-InvKey 'Sub-Total')

$r = Get-InvPdfWords -Path (Join-Path $script:InvRoot 'samples/inv_classic.pdf')
Assert-InvTrue  'the classic invoice reads'  $r.Ok
Assert-InvEqual 'one page'                   1 $r.Pages
Assert-InvTrue  'words came out'             ($r.Words.Count -gt 50)

$L = New-InvLayout -Words $r.Words -PageInfo $r.PageInfo

# a label and its value in different sizes must land on ONE line
$hit = @(Find-InvLabel -Layout $L -Labels $InvLabels.InvAmt)[0]
Assert-InvEqual 'strongest total label wins' 'total due' $hit.Phrase
$right = Get-InvRightOf -Hit $hit
Assert-InvEqual 'the total is to its right'  '950.16' $right.Text

# word boundaries: "tax" must not match inside "taxable"
$taxHits = @(Find-InvLabel -Layout $L -Labels @((New-InvLabel 'tax' 70)))
foreach ($h in $taxHits) {
    Assert-InvTrue 'tax never matched inside taxable' ($h.Line.Squash -notmatch 'taxable')
}

# a letterhead sharing a line with "INVOICE" must split into two groups
$page1 = $L.Pages[0]
$topLine = $null
foreach ($ln in $page1.Lines) { if ($ln.Text -match 'MIDWEST') { $topLine = $ln; break } }
Assert-InvTrue 'found the letterhead line' ($null -ne $topLine)
$groups = @(Split-InvLineGroups -Line $topLine)
Assert-InvEqual 'letterhead splits from the INVOICE heading' 2 $groups.Count
Assert-InvEqual 'first group is the vendor' 'MIDWEST SUPPLY CO.' $groups[0].Text
