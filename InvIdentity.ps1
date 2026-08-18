<#
================================================================================
 InvIdentity.ps1
 Knowing who WE are, so we are never reported as the vendor.

 THE PROBLEM THIS SOLVES
   Every invoice carries two companies and two addresses: the vendor's and
   ours. Nothing on the page says which is which in a machine-readable way, and
   a parser that guesses gets it wrong constantly - usually by reporting the
   customer as the vendor, which quietly corrupts a whole workpaper.

 THE FIX
   Tell the tool who we are, once (config\MyCompany.json). Then:
     1. VETO    - a name or address matching us can never be the vendor.
     2. LOCATE  - the block holding our name IS the bill-to / ship-to side.
     3. DEDUCE  - two address blocks, one is us, so the other is the vendor.
     4. REFUSE  - if the only company we can find is us, report nothing and
                  flag it. A blank beats a wrong vendor.

 SHIP-TO
   Ship-to is usually us as well, and for use tax it is often the address that
   decides the jurisdiction - goods are taxed where they are delivered, not
   where the bill is sent. So ship-to is captured whenever the block exists,
   whether or not freight is charged: plenty of invoices ship goods without a
   separate shipping line.
================================================================================
#>

Set-StrictMode -Version 2.0

<#
 Reduces a company name to a comparison key: lower case, punctuation gone,
 and the legal suffix dropped, so all of these collapse to "acmeholdings":
   Acme Holdings LLC / ACME HOLDINGS, L.L.C. / Acme Holdings, Inc.
#>
function ConvertTo-InvNameKey {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }

    $t = $Name.ToLowerInvariant()
    $t = [regex]::Replace($t, '[^a-z0-9 ]', ' ')
    $t = [regex]::Replace($t, '\s+', ' ').Trim()
    if ($t.Length -eq 0) { return '' }

    # "L.L.C." arrives as "l l c" once punctuation is stripped. Fuse a RUN of
    # single letters back into one token so the suffix list recognises it -
    # pairwise fusion is not enough, it leaves "ll c".
    $t = [regex]::Replace($t, '\b(?:[a-z]\s+){1,4}[a-z]\b', { param($m) ($m.Value -replace '\s+', '') })
    $t = [regex]::Replace($t, '\s+', ' ').Trim()

    $parts = @($t -split ' ')
    $keep = [System.Collections.ArrayList]::new()
    foreach ($p in $parts) {
        $isSuffix = $false
        foreach ($sfx in $InvCompanySuffix) {
            if ($p -eq ($sfx -replace '\.', '')) { $isSuffix = $true; break }
        }
        if (-not $isSuffix) { [void]$keep.Add($p) }
    }
    if ($keep.Count -eq 0) { $keep = $parts }        # the name was only a suffix
    return (($keep -join '') )
}

<#
 Edit distance as a 0..1 similarity. Used only as a last resort, and only at a
 high threshold - "Acme Holdings" and "Acme Holding" should match, but "Acme
 Holdings" and "Acme Plumbing" must not.
#>
function Get-InvSimilarity {
    param([string]$A, [string]$B)

    if ([string]::IsNullOrEmpty($A) -or [string]::IsNullOrEmpty($B)) { return 0.0 }
    if ($A -eq $B) { return 1.0 }

    $n = $A.Length; $m = $B.Length
    if ([Math]::Abs($n - $m) / [double][Math]::Max($n, $m) -gt 0.34) { return 0.0 }

    $prev = New-Object 'int[]' ($m + 1)
    $cur  = New-Object 'int[]' ($m + 1)
    for ($j = 0; $j -le $m; $j++) { $prev[$j] = $j }

    for ($i = 1; $i -le $n; $i++) {
        $cur[0] = $i
        for ($j = 1; $j -le $m; $j++) {
            $cost = if ($A[$i - 1] -eq $B[$j - 1]) { 0 } else { 1 }
            $d = $prev[$j] + 1
            $ins = $cur[$j - 1] + 1
            if ($ins -lt $d) { $d = $ins }
            $sub = $prev[$j - 1] + $cost
            if ($sub -lt $d) { $d = $sub }
            $cur[$j] = $d
        }
        $tmp = $prev; $prev = $cur; $cur = $tmp
    }
    $dist = $prev[$m]
    return (1.0 - ($dist / [double][Math]::Max($n, $m)))
}

<#
 Loads config\MyCompany.json and pre-computes the keys used for matching.
 Returns $null when there is no config yet (the caller runs the wizard).
#>
function Get-InvIdentity {
    [CmdletBinding()]
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $root = $PSScriptRoot
        if ([string]::IsNullOrEmpty($root)) { $root = (Get-Location).Path }
        $Path = Join-Path $root 'config\MyCompany.json'
        if (-not (Test-Path -LiteralPath $Path)) { $Path = Join-Path $root 'config/MyCompany.json' }
    }
    if (-not (Test-Path -LiteralPath $Path)) { return $null }

    $raw = $null
    try { $raw = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
    catch { throw ("config\MyCompany.json is not valid JSON: " + $_.Exception.Message) }

    $nameKeys = [System.Collections.ArrayList]::new()
    $addrKeys = [System.Collections.ArrayList]::new()
    $cityStates = [System.Collections.ArrayList]::new()

    $entities = @()
    if ($raw.PSObject.Properties.Name -contains 'Entities') { $entities = @($raw.Entities) }

    foreach ($e in $entities) {
        foreach ($nm in @($e.Name)) {
            $k = ConvertTo-InvNameKey $nm
            if ($k) { [void]$nameKeys.Add($k) }
        }
        if ($e.PSObject.Properties.Name -contains 'Aliases') {
            foreach ($al in @($e.Aliases)) {
                $k = ConvertTo-InvNameKey $al
                if ($k) { [void]$nameKeys.Add($k) }
            }
        }
        if ($e.PSObject.Properties.Name -contains 'Addresses') {
            foreach ($a in @($e.Addresses)) {
                $l1 = ''
                if ($a.PSObject.Properties.Name -contains 'Line1') { $l1 = [string]$a.Line1 }
                $k = ConvertTo-InvNameKey $l1
                if ($k) { [void]$addrKeys.Add($k) }
                $city = ''; $st = ''
                if ($a.PSObject.Properties.Name -contains 'City')  { $city = [string]$a.City }
                if ($a.PSObject.Properties.Name -contains 'State') { $st   = [string]$a.State }
                if ($city) {
                    [void]$cityStates.Add([pscustomobject]@{
                        City = $city; State = $st
                        Key = (ConvertTo-InvNameKey ($city + $st))
                    })
                }
            }
        }
    }

    $shipNames = @()
    if ($raw.PSObject.Properties.Name -contains 'ShipToNames') { $shipNames = @($raw.ShipToNames) }
    foreach ($sn in $shipNames) {
        $k = ConvertTo-InvNameKey $sn
        if ($k) { [void]$nameKeys.Add($k) }
    }

    [pscustomobject]@{
        Path        = $Path
        Raw         = $raw
        Entities    = $entities
        NameKeys    = @($nameKeys | Select-Object -Unique)
        AddressKeys = @($addrKeys | Select-Object -Unique)
        CityStates  = @($cityStates)
        AccountNumbers = $(if ($raw.PSObject.Properties.Name -contains 'AccountNumbers') { @($raw.AccountNumbers) } else { @() })
    }
}

<#
 Is this text us?

 Returns a score 0..100. 100 = an exact key match on a name or alias,
 85..99 = a strong fuzzy match, 0 = not us.
#>
function Test-InvIsSelf {
    [CmdletBinding()]
    param(
        $Identity,
        [string] $Text
    )

    if ($null -eq $Identity) { return 0 }
    if ([string]::IsNullOrWhiteSpace($Text)) { return 0 }

    $k = ConvertTo-InvNameKey $Text
    if ($k.Length -lt 3) { return 0 }

    foreach ($nk in $Identity.NameKeys) {
        if ($nk.Length -lt 3) { continue }
        if ($k -eq $nk) { return 100 }
    }
    # our name embedded in a longer line ("Sold To: Acme Holdings LLC")
    foreach ($nk in $Identity.NameKeys) {
        if ($nk.Length -lt 5) { continue }
        if ($k.Contains($nk)) { return 95 }
    }
    foreach ($nk in $Identity.NameKeys) {
        if ($nk.Length -lt 5) { continue }
        $sim = Get-InvSimilarity $k $nk
        if ($sim -ge 0.88) { return [int](80 + ($sim * 15)) }
    }
    return 0
}

<#
 Does this block of address lines belong to us? Checks the name lines and the
 street lines, so a block that names a site we own is recognised even when the
 company name is spelt differently.
#>
function Test-InvSelfBlock {
    [CmdletBinding()]
    param($Identity, $Lines)

    if ($null -eq $Identity) { return 0 }
    $best = 0
    foreach ($ln in @($Lines)) {
        $s = Test-InvIsSelf -Identity $Identity -Text $ln
        if ($s -gt $best) { $best = $s }

        $k = ConvertTo-InvNameKey $ln
        foreach ($ak in $Identity.AddressKeys) {
            if ($ak.Length -lt 5) { continue }
            if ($k -eq $ak -or $k.Contains($ak)) { if ($best -lt 92) { $best = 92 } }
        }
    }
    return $best
}

<#
 Finds the named party blocks on the invoice (Bill To, Ship To, Remit To,
 Vendor) and marks each as ours or theirs.

 Returns one entry per block found:
   Kind    BillTo | ShipTo | RemitTo | Vendor
   Lines   the text lines of the block
   IsSelf  score 0..100 that this block is us
   Page / X / Y
#>
function Get-InvPartyBlocks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Layout,
        $Identity
    )

    $out = [System.Collections.ArrayList]::new()

    # Every block heading is a stop marker for every other block, which is what
    # keeps a ship-to address from being swallowed into the bill-to block.
    $allHeadings = @()
    foreach ($kind in $InvBlockLabels.Keys) { $allHeadings += $InvBlockLabels[$kind] }

    # where every strong heading physically sits, so blocks can fence each other
    $headingHits = @(Find-InvLabel -Layout $Layout -Labels @($allHeadings | Where-Object { $_.Score -ge 85 }))

    # A block heading must be a confident one. Weak words like "customer" match
    # inside "CUSTOMER PO" and would invent a bill-to block out of a column
    # header, so they are only allowed when nothing better was found.
    $MinHeadingScore = 85

    foreach ($kind in @('BillTo','ShipTo','RemitTo','Vendor')) {
        $allHits = Find-InvLabel -Layout $Layout -Labels $InvBlockLabels[$kind]
        $hits = @($allHits | Where-Object { $_.Score -ge $MinHeadingScore })
        if ($hits.Count -eq 0) { $hits = @($allHits | Select-Object -First 1) }
        $seen = @{}

        foreach ($h in $hits) {
            if ($null -eq $h) { continue }
            # one block per heading position
            $sig = '{0}:{1}:{2}' -f $h.Page, [int]$h.Y, [int]$h.X0
            if ($seen.ContainsKey($sig)) { continue }
            $seen[$sig] = $true

            $stops = @($allHeadings | Where-Object { (ConvertTo-InvKey $_.Phrase) -ne (ConvertTo-InvKey $h.Phrase) })

            # fence this block on the right at the nearest heading to its right
            $rightBound = 0.0
            foreach ($other in $headingHits) {
                if ($other.Page -ne $h.Page) { continue }
                if ($other.X0 -le ($h.X1 + 5)) { continue }
                if ([Math]::Abs($other.Y - $h.Y) -gt 40) { continue }
                if ($rightBound -eq 0.0 -or $other.X0 -lt $rightBound) { $rightBound = $other.X0 }
            }
            if ($rightBound -gt 0) { $rightBound = $rightBound - 6 }

            # @() is mandatory: PowerShell unrolls a one-element result into a
            # bare scalar, and then .Count is gone. This is the flattening
            # footgun the address repo's ARCHITECTURE notes call out.
            $lines = @(Get-InvBlock -Layout $Layout -Hit $h -StopLabels $stops -MaxLines 6 -RightBound $rightBound)
            if ($lines.Count -eq 0) { continue }

            [void]$out.Add([pscustomobject]@{
                Kind       = $kind
                Heading    = $h.Phrase
                Score      = $h.Score
                Lines      = @($lines)
                IsSelf     = (Test-InvSelfBlock -Identity $Identity -Lines $lines)
                Page       = $h.Page
                X          = $h.X0
                Y          = $h.Y
            })
        }
    }

    return @($out)
}

<#
 Writes a starter config\MyCompany.json. Called by the first-run wizard, and
 safe to call again - it never overwrites an existing file.
#>
function New-InvIdentityFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [string] $Name,
        [string] $Line1 = '',
        [string] $City  = '',
        [string] $State = '',
        [string] $Zip   = '',
        [string[]] $Aliases = @(),
        [string[]] $ShipToNames = @()
    )

    if (Test-Path -LiteralPath $Path) { throw "That file already exists: $Path" }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $obj = [ordered]@{
        Entities = @(
            [ordered]@{
                Name      = $Name
                Aliases   = @($Aliases)
                Addresses = @(
                    [ordered]@{ Line1 = $Line1; City = $City; State = $State; Zip = $Zip; Note = 'current' }
                )
            }
        )
        ShipToNames    = @($ShipToNames)
        AccountNumbers = @()
    }

    ($obj | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $Path -Encoding UTF8
    return $Path
}
