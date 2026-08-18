<#
.SYNOPSIS
    Reads a folder of PDF invoices and writes one row per invoice to Excel.

.DESCRIPTION
    Every invoice layout is different, so nothing here assumes a template. Each
    field is found by what it sits next to, checked against strict rules, and
    cross-checked with arithmetic. Anything the parser is not sure of is left
    blank and flagged rather than guessed at.

    Pure PowerShell. Nothing is downloaded and nothing is installed: the PDF
    engine is C# source compiled at run time by the .NET Framework compiler
    already on the machine, and the .xlsx is written directly, so Excel does
    not need to be installed either.

.EXAMPLE
    .\Get-Invoices.ps1
    Asks for the folder, lets you pick columns, extracts, reviews, writes.

.EXAMPLE
    .\Get-Invoices.ps1 -Path .\PDFs -Out .\Q3.xlsx -NonInteractive
    Unattended, using the default columns.

.EXAMPLE
    .\Get-Invoices.ps1 -Path .\PDFs -Columns Vendor,InvNum,InvAmt,PONum -NoReview
#>
[CmdletBinding()]
param(
    [string]   $Path,
    [string]   $Out,
    [string[]] $Columns,
    [string]   $Preset,
    [string]   $IdentityPath,
    [string]   $SheetName = 'Invoices',
    [switch]   $Recurse,
    [switch]   $NonInteractive,
    [switch]   $NoReview
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
if ([string]::IsNullOrEmpty($root)) { $root = (Get-Location).Path }

foreach ($m in @('InvPdfText','InvDicts','InvLayout','InvValues','InvIdentity',
                 'InvFields','InvRules','InvRecord','InvColumns','InvXlsx')) {
    . (Join-Path $root ($m + '.ps1'))
}

# ------------------------------------------------------------------ presentation

function Write-InvBanner {
    Write-Host ''
    Write-Host '  PDF INVOICES -> EXCEL' -ForegroundColor Cyan
    Write-Host '  ---------------------------------------------------------------------------' -ForegroundColor DarkGray
}

function Write-InvRule { Write-Host '  ---------------------------------------------------------------------------' -ForegroundColor DarkGray }

function Read-InvAnswer {
    param([string]$Question, [string]$Default = '')
    $suffix = ''
    if ($Default -ne '') { $suffix = ' [' + $Default + ']' }
    Write-Host ('  ' + $Question + $suffix + ': ') -NoNewline -ForegroundColor Gray
    $a = Read-Host
    if ([string]::IsNullOrWhiteSpace($a)) { return $Default }
    return $a.Trim()
}

function Read-InvYesNo {
    param([string]$Question, [bool]$Default = $true)
    $d = $(if ($Default) { 'Y' } else { 'N' })
    while ($true) {
        $a = Read-InvAnswer -Question ($Question + ' (y/n)') -Default $d
        if ($a -match '^(y|yes)$') { return $true }
        if ($a -match '^(n|no)$')  { return $false }
    }
}

# ---------------------------------------------------------- first-run identity

<#
 Without this the tool cannot tell the vendor from the customer, so it is asked
 for once, before anything else, and never asked again.
#>
function Initialize-InvIdentity {
    param([string]$Root, [string]$Explicit)

    $id = $null
    try { $id = Get-InvIdentity -Path $Explicit } catch { $id = $null }
    if ($null -ne $id) { return $id }

    $target = $Explicit
    if ([string]::IsNullOrWhiteSpace($target)) { $target = Join-Path $Root 'config/MyCompany.json' }

    if ($NonInteractive) {
        throw ("There is no company profile yet, and this run is unattended.`n" +
               "Copy config\MyCompany.example.json to config\MyCompany.json and fill it in,`n" +
               "or run without -NonInteractive once to be walked through it.")
    }

    Write-Host ''
    Write-Host '  WHO ARE WE?' -ForegroundColor Cyan
    Write-InvRule
    Write-Host '  Every invoice names two companies: the vendor and us. Telling the tool who' -ForegroundColor Gray
    Write-Host '  we are is what stops our own name being written into the Vendor column, and' -ForegroundColor Gray
    Write-Host '  it is what lets it work out the vendor from the other address on the page.' -ForegroundColor Gray
    Write-Host '  Asked once. Editable later in config\MyCompany.json.' -ForegroundColor DarkGray
    Write-Host ''

    $name = ''
    while ([string]::IsNullOrWhiteSpace($name)) { $name = Read-InvAnswer -Question 'Our company name' }

    $line1 = Read-InvAnswer -Question 'Street address'
    $city  = Read-InvAnswer -Question 'City'
    $state = Read-InvAnswer -Question 'State (2 letters)'
    $zip   = Read-InvAnswer -Question 'Zip'

    Write-Host ''
    Write-Host '  Other spellings that appear on invoices we receive make the match stronger' -ForegroundColor DarkGray
    Write-Host '  (abbreviations, the old trading name, a division name). Comma separated.' -ForegroundColor DarkGray
    $aliasRaw = Read-InvAnswer -Question 'Other spellings'
    $aliases = @()
    if ($aliasRaw) { $aliases = @($aliasRaw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

    Write-Host '  Site or warehouse names that appear in Ship To blocks. Comma separated.' -ForegroundColor DarkGray
    $shipRaw = Read-InvAnswer -Question 'Ship-to names'
    $ships = @()
    if ($shipRaw) { $ships = @($shipRaw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

    [void](New-InvIdentityFile -Path $target -Name $name -Line1 $line1 -City $city `
            -State $state.ToUpperInvariant() -Zip $zip -Aliases $aliases -ShipToNames $ships)

    Write-Host ''
    Write-Host ('  Saved ' + $target) -ForegroundColor Green
    Write-Host ''

    return (Get-InvIdentity -Path $target)
}

# ------------------------------------------------------------------- row build

<#
 Turns a record into the cell values for the chosen columns, applying each
 column's normaliser and marking cells the parser was unsure about.
#>
function ConvertTo-InvRow {
    param($Record, $Cols)

    $row = @{}
    $notes = @()

    foreach ($c in @($Cols)) {
        $v = $null
        $conf = 100

        switch ($c.Key) {
            'Source'      { $v = $Record.Source; $conf = 100 }
            'Provided'    { $v = $Record.Provided; $conf = 100 }
            'Confidence'  { $v = $Record.RowConfidence; $conf = 100 }
            'NeedsReview' { $v = $(if ($Record.NeedsReview) { 'Y' } else { 'N' }); $conf = 100 }
            'Notes'       { $v = (@($Record.Issues) -join '; '); $conf = 100 }
            default {
                if ($Record.Fields.ContainsKey($c.Key)) {
                    $f = $Record.Fields[$c.Key]
                    $v = $f.Value
                    $conf = $f.Confidence
                }
            }
        }

        if ($null -ne $c.Normalize -and $null -ne $v -and ([string]$v) -ne '') {
            $v = & $c.Normalize $v
        }
        if ($c.Key -eq 'Description' -and $null -ne $v -and ([string]$v).Length -gt 200) {
            $v = ([string]$v).Substring(0, 200)
        }

        $low = ($conf -gt 0 -and $conf -lt 75) -or ($null -eq $v -and $c.Key -in @('Vendor','InvNum','InvDate','InvAmt'))
        $row[$c.Key] = @{ V = $v; Low = $low }
    }

    if ($Record.SourcePath) {
        $full = $Record.SourcePath
        try { $full = (Resolve-Path -LiteralPath $Record.SourcePath).ProviderPath } catch { }
        $row['_Link'] = $full
    }
    return $row
}

# ---------------------------------------------------------------- review pass

<#
 The review queue. Only rows the rules were not happy with, grouped by vendor
 so a layout is corrected once rather than once per invoice.
#>
function Invoke-InvReview {
    param($Records)

    $need = @($Records | Where-Object { $_.NeedsReview -and $_.Provided -eq 'Y' })
    if ($need.Count -eq 0) {
        Write-Host ''
        Write-Host '  Nothing needs review - every row passed its rules.' -ForegroundColor Green
        return
    }

    Write-Host ''
    Write-Host ('  {0} row(s) need a look.' -f $need.Count) -ForegroundColor Yellow
    if (-not (Read-InvYesNo -Question 'Review them now?' -Default $true)) { return }

    $byVendor = $need | Group-Object -Property { [string]$_.Fields['Vendor'].Value }
    $i = 0

    foreach ($grp in $byVendor) {
        foreach ($rec in $grp.Group) {
            $i++
            Clear-Host
            Write-Host ''
            Write-Host ('  {0}  |  {1}' -f $grp.Name, $rec.Source) -ForegroundColor Cyan
            Write-Host ('  row {0} of {1}   confidence {2}%' -f $i, $need.Count, $rec.RowConfidence) -ForegroundColor DarkGray
            Write-InvRule

            $editable = @('Vendor','InvNum','InvDate','InvAmt','TaxAmt','Taxable','TaxRate','Subtotal','City','State')
            $shown = [System.Collections.ArrayList]::new()
            foreach ($k in $editable) {
                if (-not $rec.Fields.ContainsKey($k)) { continue }
                $f = $rec.Fields[$k]
                $val = $f.Value
                if ($val -is [datetime]) { $val = $val.ToString('MM/dd/yyyy') }
                $colour = 'Gray'
                if ($f.Confidence -lt 60 -or $null -eq $f.Value) { $colour = 'Yellow' }
                [void]$shown.Add($k)
                Write-Host ('  {0,2}) {1,-12} {2,-28} {3}%' -f $shown.Count, $k, $val, $f.Confidence) -ForegroundColor $colour
            }

            if (@($rec.Issues).Count -gt 0) {
                Write-InvRule
                foreach ($iss in $rec.Issues) { Write-Host ('   . ' + $iss) -ForegroundColor DarkGray }
            }

            Write-InvRule
            Write-Host '  number = edit that field    t = see the page text    Enter = accept    q = stop reviewing' -ForegroundColor DarkGray
            Write-Host '  > ' -NoNewline -ForegroundColor Cyan
            $ans = Read-Host

            if ($null -eq $ans) { $ans = '' }
            $ans = $ans.Trim()
            if ($ans -eq '') { continue }
            if ($ans -match '^(q|quit)$') { return }

            if ($ans -match '^t$') {
                Show-InvPdfText -Path $rec.SourcePath -Page 1
                Write-Host ''
                Write-Host '  press Enter' -ForegroundColor DarkGray -NoNewline
                [void](Read-Host)
                continue
            }

            $idx = 0
            if ([int]::TryParse($ans, [ref]$idx) -and $idx -ge 1 -and $idx -le $shown.Count) {
                $key = $shown[$idx - 1]
                Write-Host ('  new value for ' + $key + ' (blank clears it): ') -NoNewline -ForegroundColor Gray
                $nv = Read-Host
                $parsed = $nv

                if ([string]::IsNullOrWhiteSpace($nv)) { $parsed = $null }
                elseif ($key -eq 'InvDate') {
                    $d = Convert-InvDate $nv
                    if ($null -eq $d) { Write-Host '  that is not a date I can read' -ForegroundColor Red; Start-Sleep -Milliseconds 1200; continue }
                    $parsed = $d.Date
                }
                elseif ($key -in @('InvAmt','TaxAmt','Taxable','Subtotal')) {
                    $m = Convert-InvMoney $nv
                    if ($null -eq $m) { Write-Host '  that is not an amount I can read' -ForegroundColor Red; Start-Sleep -Milliseconds 1200; continue }
                    $parsed = $m
                }
                elseif ($key -eq 'TaxRate') {
                    $p = Convert-InvPercent $nv
                    if ($null -eq $p) { $p = Convert-InvMoney $nv }
                    if ($null -eq $p) { Write-Host '  that is not a rate I can read' -ForegroundColor Red; Start-Sleep -Milliseconds 1200; continue }
                    $parsed = $p
                }

                $rec.Fields[$key] = New-InvFinding -Value $parsed -Confidence 100 -Source 'entered by hand'
                [void]$rec.Issues.Add(($key + ': corrected by hand during review'))
                $rec.RowConfidence = 100
                $rec.NeedsReview = $false
            }
        }
    }
}

# ------------------------------------------------------------------- the run

Write-InvBanner

$identity = Initialize-InvIdentity -Root $root -Explicit $IdentityPath

# --- where are the PDFs ---
if ([string]::IsNullOrWhiteSpace($Path)) {
    if ($NonInteractive) { throw 'Give me -Path: the folder holding the PDFs.' }
    $Path = Read-InvAnswer -Question 'Folder of PDFs' -Default (Get-Location).Path
}
if (-not (Test-Path -LiteralPath $Path)) { throw ("No such folder or file: " + $Path) }

$item = Get-Item -LiteralPath $Path
if ($item.PSIsContainer) {
    if (-not $NonInteractive -and -not $Recurse) {
        $Recurse = Read-InvYesNo -Question 'Include sub-folders?' -Default $false
    }
    $files = @(Get-ChildItem -LiteralPath $Path -Filter '*.pdf' -File -Recurse:$Recurse | Sort-Object FullName)
}
else {
    $files = @($item)
}

if ($files.Count -eq 0) { throw ("No PDFs found in " + $Path) }

# --- which columns ---
$colKeys = $null
if ($Columns -and @($Columns).Count -gt 0) { $colKeys = @($Columns) }
elseif ($Preset) {
    $presets = Get-InvPresets -Root $root
    if (-not $presets.ContainsKey($Preset)) { throw ("No preset called '" + $Preset + "'") }
    $colKeys = @($presets[$Preset])
}
elseif ($NonInteractive) { $colKeys = @(Get-InvDefaultColumnKeys) }
else {
    $colKeys = Show-InvColumnPicker -Root $root
    if ($null -eq $colKeys) { Write-Host '  cancelled'; return }
}
$cols = Resolve-InvColumns -Keys $colKeys

# --- extract ---
Clear-Host
Write-InvBanner
Write-Host ('  {0} PDF(s) in {1}' -f $files.Count, $item.FullName) -ForegroundColor Gray
Write-Host ('  {0} column(s): {1}' -f $cols.Count, ((@($cols | ForEach-Object { $_.Header })) -join ', ')) -ForegroundColor DarkGray
Write-InvRule
Write-Host ''

$records = [System.Collections.ArrayList]::new()
$log = [System.Collections.ArrayList]::new()
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$clean = 0; $flagged = 0; $failed = 0

for ($i = 0; $i -lt $files.Count; $i++) {
    $f = $files[$i]
    $pct = [int](100 * ($i + 1) / $files.Count)
    Write-Progress -Activity 'Reading invoices' -Status ('{0} of {1}  {2}' -f ($i + 1), $files.Count, $f.Name) -PercentComplete $pct

    $rec = $null
    try { $rec = Get-InvRecord -Path $f.FullName -Identity $identity }
    catch {
        $rec = [pscustomobject]@{
            SourcePath = $f.FullName; Source = $f.Name; Provided = 'N'; Pages = 0
            Scanned = $false; Fields = @{}
            Issues = [System.Collections.ArrayList]@(('could not be read: ' + $_.Exception.Message))
            NeedsReview = $true; RowConfidence = 0
        }
    }

    [void]$records.Add($rec)

    $state = ''
    if ($rec.Provided -ne 'Y')      { $failed++;  $state = 'no text';  $colour = 'Red' }
    elseif ($rec.NeedsReview)       { $flagged++; $state = 'review';   $colour = 'Yellow' }
    else                            { $clean++;   $state = 'ok';       $colour = 'Green' }

    $vendor = ''
    if ($rec.Fields.ContainsKey('Vendor') -and $rec.Fields['Vendor'].Value) { $vendor = [string]$rec.Fields['Vendor'].Value }
    Write-Host ('  {0,4}/{1}  ' -f ($i + 1), $files.Count) -NoNewline -ForegroundColor DarkGray
    Write-Host ('{0,-8}' -f $state) -NoNewline -ForegroundColor $colour
    Write-Host ('{0,-32} {1}' -f ($vendor.PadRight(32).Substring(0, 32)), $f.Name) -ForegroundColor Gray

    [void]$log.Add([pscustomobject]@{
        File = $f.Name
        Message = (@($rec.Issues) -join '; ')
    })
}
$sw.Stop()
Write-Progress -Activity 'Reading invoices' -Completed

Write-Host ''
Write-InvRule
Write-Host ('  {0} clean   {1} need review   {2} unreadable   ({3:N1}s)' -f `
            $clean, $flagged, $failed, $sw.Elapsed.TotalSeconds) -ForegroundColor Cyan

# --- review ---
if (-not $NonInteractive -and -not $NoReview) { Invoke-InvReview -Records $records }

# --- where does it go ---
if ([string]::IsNullOrWhiteSpace($Out)) {
    $stamp = (Get-Date).ToString('yyyy-MM-dd HHmm')
    $default = Join-Path (Get-Location).Path ('Invoices ' + $stamp + '.xlsx')
    if ($NonInteractive) { $Out = $default }
    else {
        Write-Host ''
        Write-Host '  WHERE SHOULD THIS GO?' -ForegroundColor Cyan
        Write-InvRule
        Write-Host '   1) a new workbook' -ForegroundColor Gray
        Write-Host '   2) add a sheet to a workbook I already have' -ForegroundColor Gray
        Write-Host ''
        $choice = Read-InvAnswer -Question 'Choice' -Default '1'

        if ($choice -eq '2') {
            $existing = ''
            while ($true) {
                $existing = Read-InvAnswer -Question 'Path to the existing .xlsx'
                if ($existing -and (Test-Path -LiteralPath $existing)) { break }
                Write-Host '  I cannot find that file.' -ForegroundColor Red
            }
            $Out = $existing
            $script:InvAppendMode = $true
            $SheetName = Read-InvAnswer -Question 'Name for the new sheet' -Default ('Invoices ' + (Get-Date).ToString('MM-dd'))
        }
        else {
            $Out = Read-InvAnswer -Question 'New workbook path' -Default $default
        }
    }
}

# --- write ---
$rows = [System.Collections.ArrayList]::new()
foreach ($rec in $records) { [void]$rows.Add((ConvertTo-InvRow -Record $rec -Cols $cols)) }

$csvPath = [System.IO.Path]::ChangeExtension($Out, '.csv')
$written = $null
try { $written = Export-InvCsv -Path $csvPath -Columns $cols -Rows $rows } catch { }

$appendMode = $false
if (Get-Variable -Name InvAppendMode -Scope Script -ErrorAction SilentlyContinue) { $appendMode = [bool]$script:InvAppendMode }

$xlsxPath = $null
try {
    if ($appendMode) {
        $xlsxPath = Add-InvWorksheet -Path $Out -Columns $cols -Rows $rows -SheetName $SheetName -LogLines $log
    }
    else {
        $xlsxPath = Export-InvWorkbook -Path $Out -Columns $cols -Rows $rows -LogLines $log -SheetName $SheetName
    }
}
catch {
    Write-Host ''
    Write-Host ('  The workbook could not be written: ' + $_.Exception.Message) -ForegroundColor Red
    if ($written) { Write-Host ('  The CSV is safe at ' + $written) -ForegroundColor Yellow }
    throw
}

Write-Host ''
Write-InvRule
Write-Host ('  Workbook : ' + $xlsxPath) -ForegroundColor Green
if ($written) { Write-Host ('  CSV      : ' + $written) -ForegroundColor DarkGray }
Write-Host ('  Rows     : {0}   ({1} flagged for review)' -f $rows.Count, $flagged) -ForegroundColor DarkGray
Write-Host ''
