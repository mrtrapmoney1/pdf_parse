<#
================================================================================
 InvLayout.ps1
 Words -> lines -> blocks, and the spatial searches built on them.

 WHY LINES ARE BUILT BY OVERLAP, NOT BY ROUNDING Y
   A label and its value are often set in different sizes ("TOTAL DUE:" in 11pt
   bold, "950.16" in 10pt). Rounding Y to a grid splits them onto two lines and
   the value is lost. Two words share a line when their vertical extents
   actually overlap, which is what the eye does.

 WHY TEXT IS MATCHED SQUASHED
   The same label arrives as "Invoice #:", "Invoice#", "INVOICE NO.",
   "Invoice-No" and "invoice no". Normalising to lower-case letters and digits
   only - dropping spaces, dots, dashes and colons - collapses all of those to
   "invoiceno" / "invoice#". Matches are still anchored to real word boundaries,
   so "net" cannot match inside "netamount".
================================================================================
#>

Set-StrictMode -Version 2.0

<#
 Reduces text to a comparison key: lower case, letters/digits/#/%/& only.
 "Sub-Total:" -> "subtotal"   "Invoice #" -> "invoice#"   "Bill To" -> "billto"
#>
function ConvertTo-InvKey {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $sb = [System.Text.StringBuilder]::new()
    foreach ($c in $Text.ToLowerInvariant().ToCharArray()) {
        if (($c -ge 'a' -and $c -le 'z') -or ($c -ge '0' -and $c -le '9') -or
            $c -eq '#' -or $c -eq '%' -or $c -eq '&') { [void]$sb.Append($c) }
    }
    return $sb.ToString()
}

<#
 Groups words into lines per page and indexes them for searching.

 Each line carries:
   Words   the word objects, left to right
   Text    their text joined with single spaces (what a human would read)
   Squash  the comparison key for the whole line
   Starts / Ends  char offsets in Squash where each word begins / ends
   X0,X1   left and right edge in points
   Y, YMid, Height
#>
function New-InvLayout {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [AllowEmptyCollection()] $Words,
        $PageInfo = @()
    )

    $pages = @{}
    foreach ($w in $Words) {
        if (-not $pages.ContainsKey($w.Page)) { $pages[$w.Page] = [System.Collections.ArrayList]::new() }
        [void]$pages[$w.Page].Add($w)
    }

    $pageList = [System.Collections.ArrayList]::new()

    foreach ($pn in ($pages.Keys | Sort-Object)) {
        $pw = @($pages[$pn] | Sort-Object -Property @{Expression={$_.Y}}, @{Expression={$_.X}})

        $lines = [System.Collections.ArrayList]::new()
        $cur = $null
        $curTop = 0.0; $curBot = 0.0

        foreach ($w in $pw) {
            $h = [Math]::Max(1.0, $w.H)
            $top = $w.Y
            $bot = $w.Y + $h

            $join = $false
            if ($null -ne $cur) {
                $ov = [Math]::Min($curBot, $bot) - [Math]::Max($curTop, $top)
                $minH = [Math]::Min($curBot - $curTop, $h)
                if ($minH -gt 0 -and $ov -gt ($minH * 0.45)) { $join = $true }
            }

            if (-not $join) {
                if ($null -ne $cur) { [void]$lines.Add((Close-InvLine $cur)) }
                $cur = [System.Collections.ArrayList]::new()
                $curTop = $top; $curBot = $bot
            }
            else {
                if ($top -lt $curTop) { $curTop = $top }
                if ($bot -gt $curBot) { $curBot = $bot }
            }
            [void]$cur.Add($w)
        }
        if ($null -ne $cur -and $cur.Count -gt 0) { [void]$lines.Add((Close-InvLine $cur)) }

        $info = $null
        foreach ($p in $PageInfo) { if ($p.Number -eq $pn) { $info = $p; break } }

        [void]$pageList.Add([pscustomobject]@{
            Number = $pn
            Lines  = @($lines)
            Width  = if ($info) { $info.Width }  else { 612.0 }
            Height = if ($info) { $info.Height } else { 792.0 }
        })
    }

    [pscustomobject]@{
        Pages     = @($pageList)
        PageCount = $pageList.Count
        WordCount = @($Words).Count
    }
}

function Close-InvLine {
    param($WordList)

    $ws = @($WordList | Sort-Object -Property @{Expression={$_.X}})
    $sb = [System.Text.StringBuilder]::new()
    $sq = [System.Text.StringBuilder]::new()
    $starts = [System.Collections.ArrayList]::new()
    $ends   = [System.Collections.ArrayList]::new()

    $x0 = [double]::MaxValue; $x1 = [double]::MinValue
    $top = [double]::MaxValue; $bot = [double]::MinValue

    foreach ($w in $ws) {
        if ($sb.Length -gt 0) { [void]$sb.Append(' ') }
        [void]$sb.Append($w.Text)

        $k = ConvertTo-InvKey $w.Text
        [void]$starts.Add($sq.Length)
        [void]$sq.Append($k)
        [void]$ends.Add($sq.Length)

        if ($w.X -lt $x0) { $x0 = $w.X }
        if (($w.X + $w.W) -gt $x1) { $x1 = $w.X + $w.W }
        if ($w.Y -lt $top) { $top = $w.Y }
        if (($w.Y + $w.H) -gt $bot) { $bot = $w.Y + $w.H }
    }

    [pscustomobject]@{
        Words  = $ws
        Text   = $sb.ToString()
        Squash = $sq.ToString()
        Starts = @($starts)
        Ends   = @($ends)
        X0     = $x0
        X1     = $x1
        Y      = $top
        YMid   = ($top + $bot) / 2.0
        Height = $bot - $top
        Page   = $ws[0].Page
    }
}

<#
 Finds every place a label appears.

 A hit is only accepted when it starts and ends on real word boundaries, so
 "net" never matches inside "netamount" and "tax" never matches inside
 "taxable" (which is a different field entirely).

 Returns hits sorted best-first: highest label score, then earliest page.
#>
function Find-InvLabel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Layout,
        [Parameter(Mandatory)] $Labels,
        [int] $Page = 0
    )

    $hits = [System.Collections.ArrayList]::new()

    foreach ($pg in $Layout.Pages) {
        if ($Page -gt 0 -and $pg.Number -ne $Page) { continue }

        for ($li = 0; $li -lt $pg.Lines.Count; $li++) {
            $line = $pg.Lines[$li]
            if ($line.Squash.Length -eq 0) { continue }

            foreach ($lab in $Labels) {
                $key = ConvertTo-InvKey $lab.Phrase
                if ($key.Length -eq 0) { continue }

                $from = 0
                while ($true) {
                    $idx = $line.Squash.IndexOf($key, $from, [StringComparison]::Ordinal)
                    if ($idx -lt 0) { break }
                    $from = $idx + 1

                    $wStart = -1; $wEnd = -1
                    for ($k = 0; $k -lt $line.Starts.Count; $k++) {
                        if ($line.Starts[$k] -eq $idx) { $wStart = $k }
                        if ($line.Ends[$k] -eq ($idx + $key.Length)) { $wEnd = $k }
                    }
                    if ($wStart -lt 0 -or $wEnd -lt $wStart) { continue }

                    $lw = $line.Words[$wStart]
                    $rw = $line.Words[$wEnd]

                    [void]$hits.Add([pscustomobject]@{
                        Phrase    = $lab.Phrase
                        Score     = $lab.Score
                        Page      = $pg.Number
                        LineIndex = $li
                        Line      = $line
                        WordStart = $wStart
                        WordEnd   = $wEnd
                        X0        = $lw.X
                        X1        = $rw.X + $rw.W
                        YMid      = $line.YMid
                        Y         = $line.Y
                        # a label that owns its whole line is a column header or
                        # a block heading; one with text after it is inline
                        WholeLine = ($wStart -eq 0 -and $wEnd -eq ($line.Words.Count - 1))
                    })
                }
            }
        }
    }

    return @($hits | Sort-Object -Property @{Expression={$_.Score}; Descending=$true},
                                          @{Expression={$_.Page}},
                                          @{Expression={$_.Y}})
}

<#
 The text sitting to the RIGHT of a label on the same line.
 MaxGap keeps us from reaching across a page into an unrelated column.
#>
function Get-InvRightOf {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Hit,
        [double] $MaxGap = 260.0
    )

    $line = $Hit.Line
    if ($Hit.WordEnd -ge ($line.Words.Count - 1)) { return $null }

    $rest = @($line.Words[($Hit.WordEnd + 1)..($line.Words.Count - 1)])
    if ($rest.Count -eq 0) { return $null }

    $gap = $rest[0].X - $Hit.X1
    if ($gap -gt $MaxGap) { return $null }

    [pscustomobject]@{
        Text  = (($rest | ForEach-Object { $_.Text }) -join ' ')
        Words = $rest
        Gap   = $gap
        X     = $rest[0].X
        Y     = $line.YMid
    }
}

<#
 The line(s) sitting BELOW a label, horizontally aligned with it.
 This is how "INVOICE NO." over "PE-2024-4471" is read.
#>
function Get-InvBelow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Layout,
        [Parameter(Mandatory)] $Hit,
        [int]    $MaxLines = 1,
        [double] $MaxDrop  = 46.0,
        [double] $XSlack   = 26.0
    )

    $pg = $null
    foreach ($p in $Layout.Pages) { if ($p.Number -eq $Hit.Page) { $pg = $p; break } }
    if ($null -eq $pg) { return $null }

    $out = [System.Collections.ArrayList]::new()
    $taken = 0

    for ($li = $Hit.LineIndex + 1; $li -lt $pg.Lines.Count -and $taken -lt $MaxLines; $li++) {
        $line = $pg.Lines[$li]
        if (($line.Y - $Hit.Y) -gt $MaxDrop) { break }

        $ws = @($line.Words | Where-Object {
            ($_.X + $_.W) -gt ($Hit.X0 - $XSlack) -and $_.X -lt ($Hit.X1 + $XSlack)
        })
        if ($ws.Count -eq 0) { continue }

        [void]$out.Add([pscustomobject]@{
            Text  = (($ws | ForEach-Object { $_.Text }) -join ' ')
            Words = $ws
            Y     = $line.YMid
            Drop  = $line.Y - $Hit.Y
        })
        $taken++
    }

    if ($out.Count -eq 0) { return $null }
    return @($out)
}

<#
 The block of lines under a heading such as "Bill To" or "Ship To".

 Stops at the first line that is clearly not part of the block: a vertical gap
 bigger than a line and a half, a line that starts a different column, or
 another known heading. That keeps the ship-to address out of the bill-to.
#>
function Get-InvBlock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Layout,
        [Parameter(Mandatory)] $Hit,
        [int]    $MaxLines = 7,
        [double] $XSlack   = 40.0,
        [double] $RightBound = 0.0,
        $StopLabels = @()
    )

    $pg = $null
    foreach ($p in $Layout.Pages) { if ($p.Number -eq $Hit.Page) { $pg = $p; break } }
    if ($null -eq $pg) { return @() }

    $stopKeys = @()
    foreach ($sl in $StopLabels) { $stopKeys += (ConvertTo-InvKey $sl.Phrase) }

    $left  = $Hit.X0 - $XSlack
    # Invoices set Bill To and Ship To side by side, so a block must be fenced
    # on the right at the next column - otherwise the two addresses interleave.
    $right = if ($RightBound -gt 0) { $RightBound } else { $Hit.X1 + 220.0 }
    $out   = [System.Collections.ArrayList]::new()
    $prevY = $Hit.Y
    $lineH = [Math]::Max(8.0, $Hit.Line.Height)

    # the heading may sit on the same line as the first value ("Bill To: Acme")
    if (-not $Hit.WholeLine) {
        $r = Get-InvRightOf -Hit $Hit -MaxGap 200
        if ($r) { [void]$out.Add($r.Text) }
    }

    for ($li = $Hit.LineIndex + 1; $li -lt $pg.Lines.Count -and $out.Count -lt $MaxLines; $li++) {
        $line = $pg.Lines[$li]

        if (($line.Y - $prevY) -gt ($lineH * 2.6)) { break }        # blank run

        $ws = @($line.Words | Where-Object { $_.X -ge $left -and $_.X -lt $right })
        if ($ws.Count -eq 0) {
            if (($line.Y - $prevY) -gt ($lineH * 1.8)) { break }
            continue
        }

        $txt = (($ws | ForEach-Object { $_.Text }) -join ' ')
        $k = ConvertTo-InvKey $txt
        $isStop = $false
        foreach ($sk in $stopKeys) {
            if ($sk.Length -gt 0 -and $k.StartsWith($sk, [StringComparison]::Ordinal)) { $isStop = $true; break }
        }
        if ($isStop) { break }

        [void]$out.Add($txt)
        $prevY = $line.Y
    }

    return @($out)
}

<#
 Every line on a page, top to bottom, as plain text. Used for whole-document
 scans (finding a date anywhere, spotting "page 2 of 2") and for the review
 screen's context display.
#>
function Get-InvPageLines {
    param([Parameter(Mandatory)] $Layout, [int]$Page = 0)
    $out = [System.Collections.ArrayList]::new()
    foreach ($pg in $Layout.Pages) {
        if ($Page -gt 0 -and $pg.Number -ne $Page) { continue }
        foreach ($l in $pg.Lines) { [void]$out.Add($l) }
    }
    return @($out)
}
