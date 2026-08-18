<#
================================================================================
 InvFields.ps1
 Finding the actual values. One extractor per field, each returning a value,
 a confidence, and where it came from.

 THE PRINCIPLE: ACCURACY OVER SPEED
   Every field is attacked from several directions - the label to its left, the
   label above it, the column it sits in, and arithmetic. When two independent
   strategies agree, confidence goes UP. When they disagree, the field is
   reported with a conflict note and low confidence rather than a coin flip.
   When nothing is solid, the field comes back empty with a reason.

 Nothing here writes anything or asks anything. It reads a layout and returns
 findings, so it can be tested on its own.
================================================================================
#>

Set-StrictMode -Version 2.0

function New-InvFinding {
    param(
        $Value,
        [int]    $Confidence = 0,
        [string] $Source     = '',
        [string] $Note       = '',
        [string] $Raw        = '',
        [int]    $Page       = 0
    )
    [pscustomobject]@{
        Value = $Value; Confidence = $Confidence; Source = $Source
        Note  = $Note;  Raw = $Raw; Page = $Page
    }
}

# ------------------------------------------------------------------ addresses

<#
 Pulls City / State / Zip / Street out of a block of address lines.

 Works from the bottom up, because the city/state/zip line is always the last
 real line of a US address. Accepts "Omaha, NE 68102", "Omaha NE 68102" and
 "OMAHA, NEBRASKA 68102-1234".
#>
function Get-InvAddressParts {
    [CmdletBinding()]
    param([AllowEmptyCollection()] $Lines)

    $res = [pscustomobject]@{
        Name = ''; Street = ''; City = ''; State = ''; Zip = ''
        CityLineIndex = -1; Confidence = 0
    }
    $ls = @($Lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($ls.Count -eq 0) { return $res }

    # Work from the bottom: the city/state/zip is always the last real line of a
    # US address. It is often NOT on a line of its own though - "PO Box 4410,
    # Kearney, NE 68848" and "1201 W Adams | Omaha, NE 68132" are both common -
    # so anchor on the STATE + ZIP at the end of the line and read backwards.
    for ($i = $ls.Count - 1; $i -ge 0; $i--) {
        $t = ([string]$ls[$i]).Trim().TrimEnd(',', '.', ';')
        if ($t.Length -lt 5) { continue }

        $st = ''; $zip = ''; $head = ''

        $m = [regex]::Match($t, '(?<st>[A-Za-z]{2})\.?\s+(?<zip>\d{5}(?:-\d{4})?)\s*$')
        if ($m.Success -and $InvStateSet.ContainsKey($m.Groups['st'].Value.ToUpperInvariant())) {
            $st  = $m.Groups['st'].Value.ToUpperInvariant()
            $zip = $m.Groups['zip'].Value
            $head = $t.Substring(0, $m.Index)
        }
        else {
            # state spelled out: "Omaha, Nebraska 68102"
            $m = [regex]::Match($t, '(?<st>[A-Za-z]{4,20})\s+(?<zip>\d{5}(?:-\d{4})?)\s*$')
            if ($m.Success -and $InvStateAbbr.ContainsKey($m.Groups['st'].Value.ToLowerInvariant())) {
                $st  = $InvStateAbbr[$m.Groups['st'].Value.ToLowerInvariant()]
                $zip = $m.Groups['zip'].Value
                $head = $t.Substring(0, $m.Index)
            }
            else {
                # no zip at all: "Omaha, NE"
                $m = [regex]::Match($t, ',\s*(?<st>[A-Za-z]{2})\.?\s*$')
                if ($m.Success -and $InvStateSet.ContainsKey($m.Groups['st'].Value.ToUpperInvariant())) {
                    $st = $m.Groups['st'].Value.ToUpperInvariant()
                    $head = $t.Substring(0, $m.Index)
                }
            }
        }

        if ($st -eq '') { continue }

        $head = $head.TrimEnd(' ', ',', '|', '-', ';')
        $city = ''

        # Prefer the segment after the last comma or pipe - that is the city in
        # every conventionally punctuated address.
        $segs = @($head -split '[,|]')
        $tail = ''
        if ($segs.Count -gt 0) { $tail = ([string]$segs[$segs.Count - 1]).Trim() }

        if ($tail -match '^[A-Za-z][A-Za-z .''\-]{1,40}$' -and $tail.Length -ge 2) {
            $city = $tail
        }
        else {
            # Unpunctuated: walk back over whole words, stopping at a number or
            # a street type, so "4820 South 72nd Street Lincoln" gives "Lincoln"
            # and not "Street Lincoln".
            $words = @($head -split '\s+' | Where-Object { $_ -ne '' })
            $take = [System.Collections.ArrayList]::new()
            for ($k = $words.Count - 1; $k -ge 0 -and $take.Count -lt 3; $k--) {
                $wd = ([string]$words[$k]).Trim(',', '.', '|')
                if ($wd -match '\d') { break }
                if ($wd -notmatch '^[A-Za-z][A-Za-z.''\-]*$') { break }
                if ($InvStreetType.ContainsKey($wd.ToLowerInvariant())) { break }
                [void]$take.Insert(0, $wd)
            }
            if ($take.Count -gt 0) { $city = ($take -join ' ') }
        }

        if ($city -eq '') { continue }

        $res.City  = (Get-Culture).TextInfo.ToTitleCase($city.Trim().ToLowerInvariant())
        $res.State = $st
        $res.Zip   = $zip
        $res.CityLineIndex = $i
        $res.Confidence = $(if ($zip) { 95 } else { 80 })
        break
    }

    if ($ls.Count -gt 0) { $res.Name = ([string]$ls[0]).Trim() }

    # the street is whatever sits between the name and the city line
    if ($res.CityLineIndex -gt 0) {
        $streetLines = @()
        for ($i = 1; $i -lt $res.CityLineIndex; $i++) { $streetLines += ([string]$ls[$i]).Trim() }
        if ($streetLines.Count -eq 0 -and $res.CityLineIndex -eq 1) { $streetLines = @() }
        $res.Street = ($streetLines -join ', ')
    }
    elseif ($ls.Count -gt 1) {
        $res.Street = ([string]$ls[1]).Trim()
    }

    return $res
}

<#
 Does this line look like a company name rather than a street, a phone number
 or a document title?
#>
function Test-InvCompanyLine {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    $t = $Text.Trim()

    if ($t.Length -lt 3 -or $t.Length -gt 70) { return $false }
    if ($InvNotAName.ContainsKey($t.ToLowerInvariant())) { return $false }
    if ($t -match '^\d') { return $false }                                  # starts with a house number
    if ($t -match '^(?:p\.?\s?o\.?\s+box)\b') { return $false }
    if ($t -match '^\(?\d{3}\)?[-. ]\d{3}[-. ]\d{4}') { return $false }     # phone
    if ($t -match '@') { return $false }                                    # email
    if ($t -match '^(?:www\.|https?://)') { return $false }
    if ($t -notmatch '[A-Za-z]') { return $false }

    # a line ending in a street type is an address, not a name
    $last = ($t.TrimEnd('.', ',') -split '\s+')[-1]
    if ($InvStreetType.ContainsKey($last.ToLowerInvariant())) { return $false }

    # a line that ends in a colon is a label introducing a value, not a name
    if ($t.TrimEnd() -match ':$') { return $false }

    # A line that IS an address line is not a name. "Omaha, NE 68102" is the
    # tail of somebody's address, not a company, and a blacklist of bad words
    # will never catch it - so test structurally instead.
    $asAddr = Get-InvAddressParts -Lines @($t)
    if ($asAddr.State -and $asAddr.CityLineIndex -eq 0) { return $false }
    if ($t -match '^\s*\d') { return $false }
    if ($t -match '\b\d{5}(-\d{4})?\s*$') { return $false }

    # a document title, however large it is printed
    $k = ConvertTo-InvKey $t
    foreach ($bad in $InvNotAName.Keys) {
        if ($k -eq (ConvertTo-InvKey $bad)) { return $false }
    }

    # A FIELD LABEL is not a company. This matters most when our own name has
    # just been vetoed as the vendor: without it the letterhead fallback moves
    # on to the next big line and happily reports "Invoice #:" as the vendor.
    # Matched EXACTLY, so a real vendor called "Total Wine" still survives.
    foreach ($set in @($InvLabels, $InvBlockLabels)) {
        foreach ($field in $set.Keys) {
            foreach ($lab in $set[$field]) {
                if ($lab.Score -lt 60) { continue }
                if ($k -eq (ConvertTo-InvKey $lab.Phrase)) { return $false }
            }
        }
    }

    return $true
}

# --------------------------------------------------------------------- vendor

<#
 Works out who billed us.

 Strategies, most trusted first:
   A. a block explicitly headed "Sold By" / "Vendor" / "Remit To" that is not us
   B. the letterhead - the largest text at the top of page 1 that is a company
      name, is not us, and is not the word "INVOICE"
   C. the only party block that is not us

 Agreement between A/C and B pushes confidence to the top. If the only company
 we can find is ourselves, this returns nothing rather than guessing.
#>
function Get-InvVendor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Layout,
        $Identity,
        $Blocks = @()
    )

    $cands = [System.Collections.ArrayList]::new()
    $selfLetterhead = $false

    # --- A / C: explicit blocks -------------------------------------------
    foreach ($b in @($Blocks)) {
        if ($b.IsSelf -ge 85) { continue }
        $parts = Get-InvAddressParts -Lines $b.Lines
        if (-not (Test-InvCompanyLine $parts.Name)) { continue }

        $conf = 0
        if ($b.Kind -eq 'Vendor')  { $conf = 92 }
        elseif ($b.Kind -eq 'RemitTo') { $conf = 86 }
        else { continue }          # a non-self BillTo/ShipTo is not the vendor

        [void]$cands.Add([pscustomobject]@{
            Name = $parts.Name; Parts = $parts; Conf = $conf
            Source = ('block:' + $b.Kind); Lines = $b.Lines
        })
    }

    # --- B: the letterhead -------------------------------------------------
    $page1 = $null
    foreach ($p in $Layout.Pages) { if ($p.Number -eq 1) { $page1 = $p; break } }

    if ($null -ne $page1) {
        $topCut = $page1.Height * 0.34
        $top = @($page1.Lines | Where-Object { $_.Y -lt $topCut })

        # Rank every candidate TWICE: the strongest overall, and the strongest
        # that is not us. If the strongest overall IS us, this document is our
        # own letterhead - a credit memo we issued, or something misfiled - and
        # the honest answer is that there is no vendor on it. Falling through to
        # the next-biggest line is how a parser ends up reporting a field label
        # or a city as the vendor.
        $best = $null
        $bestAny = $null
        foreach ($ln in $top) {
            # The letterhead shares its line with the word "INVOICE" set far to
            # the right. Judge each column group on its own, never the whole line.
            foreach ($g in @(Split-InvLineGroups -Line $ln)) {
                $t = $g.Text.Trim()
                if (-not (Test-InvCompanyLine $t)) { continue }

                # bigger and higher wins; bold helps
                $rank = $g.MaxSize * 10 + $(if ($g.Bold) { 6 } else { 0 }) - ($ln.Y / 40.0)
                $cand = [pscustomobject]@{
                    Line = $ln; Group = $g; Text = $t; Rank = $rank; Size = $g.MaxSize
                }

                if ($null -eq $bestAny -or $rank -gt $bestAny.Rank) { $bestAny = $cand }

                if ($null -ne $Identity -and (Test-InvIsSelf -Identity $Identity -Text $t) -ge 85) { continue }
                if ($null -eq $best -or $rank -gt $best.Rank) { $best = $cand }
            }
        }

        if ($null -ne $bestAny -and $null -ne $Identity -and
            (Test-InvIsSelf -Identity $Identity -Text $bestAny.Text) -ge 85) {
            $selfLetterhead = $true
            $best = $null
        }

        if ($null -ne $best) {
            # the address underneath the letterhead belongs to it
            $addrLines = @($best.Text)
            $idx = [Array]::IndexOf($page1.Lines, $best.Line)
            if ($idx -ge 0) {
                $prevY = $best.Line.Y
                # stay inside the letterhead's own column, or the invoice-number
                # block on the right gets swept into the vendor's address
                $colL = $best.Group.X0 - 40
                $colR = $best.Group.X1 + 90
                for ($i = $idx + 1; $i -lt $page1.Lines.Count -and $addrLines.Count -lt 5; $i++) {
                    $nl = $page1.Lines[$i]
                    if (($nl.Y - $prevY) -gt ($best.Line.Height * 3.0)) { break }
                    $ws = @($nl.Words | Where-Object { $_.X -lt $colR -and ($_.X + $_.W) -gt $colL })
                    if ($ws.Count -eq 0) { break }
                    $txt = (($ws | ForEach-Object { $_.Text }) -join ' ')
                    if ($null -ne $Identity -and (Test-InvIsSelf -Identity $Identity -Text $txt) -ge 85) { break }
                    $addrLines += $txt
                    $prevY = $nl.Y
                }
            }

            $parts = Get-InvAddressParts -Lines $addrLines
            $conf = 74
            if ($best.Size -ge 13) { $conf += 6 }
            [void]$cands.Add([pscustomobject]@{
                Name = $best.Text; Parts = $parts; Conf = $conf
                Source = 'letterhead'; Lines = $addrLines
            })
        }
    }

    if ($cands.Count -eq 0) {
        $why = 'No vendor could be identified on the page'
        if ($selfLetterhead) {
            $why = 'The letterhead on this document is our own company, so there is no vendor on it - it may be a credit memo we issued, or a file that does not belong in this folder'
        }
        return [pscustomobject]@{
            Name = New-InvFinding -Value $null -Confidence 0 -Note $why
            Address = New-InvFinding -Value $null -Confidence 0
            City = New-InvFinding -Value $null -Confidence 0
            State = New-InvFinding -Value $null -Confidence 0
            Zip = New-InvFinding -Value $null -Confidence 0
        }
    }

    $ranked = @($cands | Sort-Object -Property @{Expression={$_.Conf}; Descending=$true})
    $win = $ranked[0]
    $conf = $win.Conf
    $note = ''

    # independent agreement is the strongest signal we have
    if ($ranked.Count -gt 1) {
        $k0 = ConvertTo-InvNameKey $win.Name
        $agree = $false
        foreach ($o in $ranked[1..($ranked.Count - 1)]) {
            $k1 = ConvertTo-InvNameKey $o.Name
            if ($k0 -eq $k1 -or (Get-InvSimilarity $k0 $k1) -ge 0.9) { $agree = $true; break }
        }
        if ($agree) { $conf = [Math]::Min(99, $conf + 14) }
        else { $note = 'Two different vendor names were found; the stronger one was used' ; $conf -= 8 }
    }

    $parts = $win.Parts
    $addrText = @($win.Lines | Select-Object -Skip 1) -join ', '

    [pscustomobject]@{
        Name    = New-InvFinding -Value $win.Name -Confidence $conf -Source $win.Source -Note $note
        Address = New-InvFinding -Value $addrText -Confidence $(if ($addrText) { $conf - 6 } else { 0 }) -Source $win.Source
        City    = New-InvFinding -Value $parts.City  -Confidence $(if ($parts.City)  { [Math]::Min($conf, $parts.Confidence) } else { 0 }) -Source $win.Source
        State   = New-InvFinding -Value $parts.State -Confidence $(if ($parts.State) { [Math]::Min($conf, $parts.Confidence) } else { 0 }) -Source $win.Source
        Zip     = New-InvFinding -Value $parts.Zip   -Confidence $(if ($parts.Zip)   { [Math]::Min($conf, $parts.Confidence) } else { 0 }) -Source $win.Source
    }
}

# ---------------------------------------------------------------- generic get

<#
 Reads a money field by its label.

 Takes the RIGHTMOST money value on the label's line, which is what makes
 "NE Sales Tax 7.0%: 62.16" yield 62.16 rather than 7. Falls back to the line
 below when the label stands alone (labels-above-values layouts).
#>
function Get-InvMoneyField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Layout,
        [Parameter(Mandatory)] $Labels,
        [int] $MaxHits = 6,
        [switch] $PreferLastPage
    )

    $hits = @(Find-InvLabel -Layout $Layout -Labels $Labels)
    if ($hits.Count -eq 0) { return @() }

    $out = [System.Collections.ArrayList]::new()
    $n = 0

    foreach ($h in $hits) {
        if ($n -ge $MaxHits) { break }
        $n++

        $conf = $h.Score
        $val = $null
        $raw = ''
        $src = 'right of "' + $h.Phrase + '"'

        $r = Get-InvRightOf -Hit $h
        if ($null -ne $r) {
            $toks = @(Get-InvMoneyTokens $r.Text)
            if ($toks.Count -gt 0) {
                $val = $toks[$toks.Count - 1].Value
                $raw = $toks[$toks.Count - 1].Text
                if ($toks.Count -gt 1) { $conf -= 4 }
                if ($r.Gap -gt 200) { $conf -= 6 }
            }
        }

        if ($null -eq $val) {
            $below = @(Get-InvBelow -Layout $Layout -Hit $h -MaxLines 1)
            if ($below.Count -gt 0) {
                $toks = @(Get-InvMoneyTokens $below[0].Text)
                if ($toks.Count -gt 0) {
                    $val = $toks[0].Value
                    $raw = $toks[0].Text
                    $conf -= 6
                    $src = 'below "' + $h.Phrase + '"'
                }
            }
        }

        if ($null -eq $val) { continue }

        if ($PreferLastPage -and $h.Page -lt $Layout.PageCount) { $conf -= 10 }

        [void]$out.Add((New-InvFinding -Value $val -Confidence ([Math]::Max(1, $conf)) `
                        -Source $src -Raw $raw -Page $h.Page))
    }

    return @($out)
}

<#
 Collapses several findings for one field into a single answer.

 Independent agreement raises confidence; disagreement lowers it and records
 why, so the review screen can show the conflict instead of hiding it.
#>
function Resolve-InvFinding {
    [CmdletBinding()]
    param([AllowEmptyCollection()] $Findings, [double] $Tolerance = 0.005)

    $fs = @($Findings | Where-Object { $null -ne $_ -and $null -ne $_.Value })
    if ($fs.Count -eq 0) { return (New-InvFinding -Value $null -Confidence 0) }

    $ranked = @($fs | Sort-Object -Property @{Expression={$_.Confidence}; Descending=$true})
    $win = $ranked[0]
    if ($ranked.Count -eq 1) { return $win }

    $same = 0; $diff = 0
    foreach ($o in $ranked[1..($ranked.Count - 1)]) {
        $isSame = $false
        if ($win.Value -is [double] -and $o.Value -is [double]) {
            $isSame = ([Math]::Abs($win.Value - $o.Value) -le ([Math]::Max(0.01, [Math]::Abs($win.Value) * $Tolerance)))
        }
        elseif ($win.Value -is [datetime] -and $o.Value -is [datetime]) {
            $isSame = ($win.Value -eq $o.Value)
        }
        else {
            $isSame = ([string]$win.Value -eq [string]$o.Value)
        }
        if ($isSame) { $same++ } else { $diff++ }
    }

    $conf = $win.Confidence
    $note = $win.Note
    if ($same -gt 0) { $conf = [Math]::Min(99, $conf + 8) }
    if ($diff -gt 0) {
        $conf = [Math]::Max(1, $conf - 12)
        $others = (@($ranked[1..($ranked.Count-1)] | ForEach-Object { $_.Value }) -join ', ')
        $note = (($note + ' ').Trim() + ' Other candidates found: ' + $others).Trim()
    }

    New-InvFinding -Value $win.Value -Confidence $conf -Source $win.Source -Note $note -Raw $win.Raw -Page $win.Page
}

# ------------------------------------------------------------- text and dates

<#
 Reads a text field by its label, validating with the supplied test before
 accepting. Without the validator this would happily return the next label on
 the line as if it were a value.
#>
function Get-InvTextField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Layout,
        [Parameter(Mandatory)] $Labels,
        [scriptblock] $Validator = $null,
        [int] $MaxHits = 6,
        [int] $MaxWords = 6
    )

    $hits = @(Find-InvLabel -Layout $Layout -Labels $Labels)
    $out = [System.Collections.ArrayList]::new()
    $n = 0

    foreach ($h in $hits) {
        if ($n -ge $MaxHits) { break }
        $n++

        $tries = [System.Collections.ArrayList]::new()

        $r = Get-InvRightOf -Hit $h -MaxGap 200
        if ($null -ne $r) {
            $words = @($r.Words | Select-Object -First $MaxWords)
            # The first word alone is tried FIRST: an invoice number is one
            # token, and reaching further right only picks up the next column.
            if ($words.Count -gt 0) {
                [void]$tries.Add([pscustomobject]@{
                    Text = $words[0].Text; Conf = $h.Score
                    Src = 'right of "' + $h.Phrase + '"'
                })
            }
            if ($words.Count -gt 1) {
                [void]$tries.Add([pscustomobject]@{
                    Text = (($words | ForEach-Object { $_.Text }) -join ' ')
                    Conf = $h.Score - 2
                    Src  = 'right of "' + $h.Phrase + '"'
                })
            }
        }

        $below = @(Get-InvBelow -Layout $Layout -Hit $h -MaxLines 1)
        if ($below.Count -gt 0) {
            $bw = @($below[0].Words | Select-Object -First $MaxWords)
            [void]$tries.Add([pscustomobject]@{
                Text = (($bw | ForEach-Object { $_.Text }) -join ' ')
                Conf = $h.Score - 6
                Src  = 'below "' + $h.Phrase + '"'
            })
            if ($bw.Count -gt 1) {
                [void]$tries.Add([pscustomobject]@{
                    Text = $bw[0].Text; Conf = $h.Score - 8
                    Src = 'below "' + $h.Phrase + '"'
                })
            }
        }

        foreach ($t in $tries) {
            $txt = ([string]$t.Text).Trim().TrimEnd(':', ',', ';')
            if ([string]::IsNullOrWhiteSpace($txt)) { continue }
            if ($null -ne $Validator) {
                $ok = & $Validator $txt
                if (-not $ok) { continue }
            }
            [void]$out.Add((New-InvFinding -Value $txt -Confidence ([Math]::Max(1, $t.Conf)) `
                            -Source $t.Src -Raw $txt -Page $h.Page))
            break                                   # first accepted try per hit
        }
    }

    return @($out)
}

<#
 Reads a date field. Ambiguous day/month forms are accepted but flagged, never
 silently resolved.
#>
function Get-InvDateField {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Layout, [Parameter(Mandatory)] $Labels, [int]$MaxHits = 6)

    $hits = @(Find-InvLabel -Layout $Layout -Labels $Labels)
    $out = [System.Collections.ArrayList]::new()
    $n = 0

    foreach ($h in $hits) {
        if ($n -ge $MaxHits) { break }
        $n++

        $texts = [System.Collections.ArrayList]::new()
        $r = Get-InvRightOf -Hit $h -MaxGap 200
        if ($null -ne $r) {
            [void]$texts.Add([pscustomobject]@{ T = $r.Text; C = $h.Score; S = 'right of "' + $h.Phrase + '"' })
            # "April 2, 2024" is three words; also try progressively shorter runs
            $ws = @($r.Words)
            for ($take = [Math]::Min(4, $ws.Count); $take -ge 1; $take--) {
                $t = (($ws | Select-Object -First $take | ForEach-Object { $_.Text }) -join ' ')
                [void]$texts.Add([pscustomobject]@{ T = $t; C = $h.Score - 1; S = 'right of "' + $h.Phrase + '"' })
            }
        }
        $below = @(Get-InvBelow -Layout $Layout -Hit $h -MaxLines 1)
        if ($below.Count -gt 0) {
            $ws = @($below[0].Words)
            for ($take = [Math]::Min(4, $ws.Count); $take -ge 1; $take--) {
                $t = (($ws | Select-Object -First $take | ForEach-Object { $_.Text }) -join ' ')
                [void]$texts.Add([pscustomobject]@{ T = $t; C = $h.Score - 6; S = 'below "' + $h.Phrase + '"' })
            }
        }

        foreach ($cand in $texts) {
            $d = Convert-InvDate ([string]$cand.T)
            if ($null -eq $d) { continue }
            $note = ''
            $conf = $cand.C
            if ($d.Ambiguous) {
                $note = 'Day/month order is ambiguous on this invoice; read as US month/day'
                $conf -= 10
            }
            [void]$out.Add((New-InvFinding -Value $d.Date -Confidence ([Math]::Max(1, $conf)) `
                            -Source $cand.S -Note $note -Raw ([string]$cand.T) -Page $h.Page))
            break
        }
    }

    return @($out)
}

function Get-InvPercentField {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Layout, [Parameter(Mandatory)] $Labels, [int]$MaxHits = 5)

    $hits = @(Find-InvLabel -Layout $Layout -Labels $Labels)
    $out = [System.Collections.ArrayList]::new()
    $n = 0

    foreach ($h in $hits) {
        if ($n -ge $MaxHits) { break }
        $n++
        $texts = @()
        $r = Get-InvRightOf -Hit $h -MaxGap 200
        if ($null -ne $r) { $texts += ,@($r.Text, $h.Score, ('right of "' + $h.Phrase + '"')) }
        $below = @(Get-InvBelow -Layout $Layout -Hit $h -MaxLines 1)
        if ($below.Count -gt 0) { $texts += ,@($below[0].Text, ($h.Score - 6), ('below "' + $h.Phrase + '"')) }

        foreach ($t in $texts) {
            $m = [regex]::Match([string]$t[0], '\d{1,3}(?:\.\d{1,4})?\s*%')
            if (-not $m.Success) { continue }
            $p = Convert-InvPercent $m.Value
            if ($null -eq $p) { continue }
            [void]$out.Add((New-InvFinding -Value $p -Confidence ([Math]::Max(1, [int]$t[1])) `
                            -Source ([string]$t[2]) -Raw $m.Value -Page $h.Page))
            break
        }
    }
    return @($out)
}

<#
 Picks a description of what was bought.

 First choice is the first real cell under a "Description" column header. If
 there is no such table, the fallback is the longest prose line in the body of
 page 1 that is not an address, a label row or a totals row. Confidence stays
 modest either way - a description is a summary, not a fact.
#>
function Get-InvDescription {
    [CmdletBinding()]
    param([Parameter(Mandatory)] $Layout, $Identity)

    $descLabels = @(
        (New-InvLabel 'description'      95), (New-InvLabel 'item description' 96)
        (New-InvLabel 'product'          80), (New-InvLabel 'service'          80)
        (New-InvLabel 'details'          78), (New-InvLabel 'item'             70)
    )

    $hits = @(Find-InvLabel -Layout $Layout -Labels $descLabels)
    foreach ($h in $hits) {
        # It must be a COLUMN HEADER, not the word "service" inside a sentence.
        # A header starts its line, has few words, AND is short in characters -
        # "Service performed: emergency panel repair and breaker replacement" is
        # eight words too, and without the length test it passes as a header.
        if ($h.WordStart -ne 0) { continue }
        if ($h.Line.Words.Count -gt 8) { continue }
        if ($h.Line.Text.Length -gt 45 -and -not $h.WholeLine) { continue }

        # The value spans the whole description COLUMN, not just the width of
        # the word "Description" - otherwise "Galvanized conduit 3/4 in." comes
        # back as "Galvanized conduit". Fence it at the next column header.
        $right = 0.0
        foreach ($g in @(Split-InvLineGroups -Line $h.Line)) {
            if ($g.X0 -gt ($h.X1 + 4)) { $right = $g.X0 - 4; break }
        }
        if ($right -le 0) { $right = $h.X0 + 320 }

        $pg = $null
        foreach ($p in $Layout.Pages) { if ($p.Number -eq $h.Page) { $pg = $p; break } }
        if ($null -eq $pg) { continue }

        $found = $null
        for ($li = $h.LineIndex + 1; $li -lt $pg.Lines.Count -and $li -le ($h.LineIndex + 4); $li++) {
            $ln = $pg.Lines[$li]
            if (($ln.Y - $h.Y) -gt 70) { break }
            $ws = @($ln.Words | Where-Object { $_.X -ge ($h.X0 - 8) -and $_.X -lt $right })
            if ($ws.Count -eq 0) { continue }
            $t = (($ws | ForEach-Object { $_.Text }) -join ' ').Trim()
            if ($t.Length -lt 4) { continue }
            if ($t -match '^[\d\s.,$%()-]+$') { continue }        # all numbers
            $found = $t
            break
        }

        if ($found) {
            return (New-InvFinding -Value $found -Confidence ([Math]::Min(88, $h.Score - 6)) `
                    -Source 'description column' -Page $h.Page)
        }
    }

    # fallback: the longest piece of prose in the body of page 1
    $page1 = $null
    foreach ($p in $Layout.Pages) { if ($p.Number -eq 1) { $page1 = $p; break } }
    if ($null -eq $page1) { return (New-InvFinding -Value $null -Confidence 0) }

    # Prose often runs over two or three lines. Judge BLOCKS of consecutive
    # prose, not single lines, or a wrapped sentence loses to its own tail.
    $best = $null
    $run = ''
    $runPrevY = -999.0
    $runH = 10.0
    $runLines = 0

    foreach ($ln in $page1.Lines) {
        $t = $ln.Text.Trim()
        $isProse = $true
        if ($ln.Y -lt ($page1.Height * 0.22) -or $ln.Y -gt ($page1.Height * 0.75)) { $isProse = $false }
        if ($isProse -and ($t.Length -lt 12 -or $t.Length -gt 160)) { $isProse = $false }

        if ($isProse) {
            $alpha = @([regex]::Matches($t, '[A-Za-z]{2,}')).Count
            if ($alpha -lt 3) { $isProse = $false }
            elseif ($t -match '\d{5}(-\d{4})?\s*$') { $isProse = $false }     # an address
            elseif ($null -ne $Identity -and (Test-InvIsSelf -Identity $Identity -Text $t) -ge 85) { $isProse = $false }
            else {
                $k = ConvertTo-InvKey $t
                foreach ($fld in @('InvAmt','Subtotal','TaxAmt','TaxRate','TaxableAmount')) {
                    foreach ($lab in $InvLabels[$fld]) {
                        if ($lab.Score -lt 85) { continue }
                        if ($k.StartsWith((ConvertTo-InvKey $lab.Phrase), [StringComparison]::Ordinal)) { $isProse = $false; break }
                    }
                    if (-not $isProse) { break }
                }
            }
        }

        if (-not $isProse) {
            if ($run -and ($null -eq $best -or $run.Length -gt $best.Length)) { $best = $run }
            $run = ''
            $runLines = 0
            $runPrevY = -999.0
            continue
        }

        # strip the amount that sits at the end of a line-item row before joining
        $tClean = [regex]::Replace($t, '\s+\(?\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})\)?$', '').Trim()

        if ($run -and $runLines -lt 3 -and (($ln.Y - $runPrevY) -le ($runH * 2.2))) {
            $run = $run + ' ' + $tClean
            $runLines++
        }
        else {
            if ($run -and ($null -eq $best -or $run.Length -gt $best.Length)) { $best = $run }
            $run = $tClean
            $runLines = 1
        }
        $runPrevY = $ln.Y
        $runH = [Math]::Max(6.0, $ln.Height)
    }
    if ($run -and ($null -eq $best -or $run.Length -gt $best.Length)) { $best = $run }

    if ($null -eq $best) { return (New-InvFinding -Value $null -Confidence 0 -Note 'No description text found') }

    # a trailing amount belongs to the money column, not to the description
    $best = [regex]::Replace($best, '\s+\(?\$?\d{1,3}(?:,\d{3})*(?:\.\d{2})\)?$', '').Trim()
    if ($best.Length -gt 200) { $best = $best.Substring(0, 200).Trim() }

    return (New-InvFinding -Value $best -Confidence 58 -Source 'body text' -Page 1)
}
