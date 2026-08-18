<#
================================================================================
 InvPdfText.ps1
 PDF -> positioned words. The foundation everything else stands on.

 USE
   . .\InvPdfText.ps1
   $r = Get-InvPdfWords -Path .\invoice.pdf
   $r.Words[0]        # Page, Text, X, Y, W, H, Size, Bold  (points, top-left)
   $r.Ok / $r.Note

   Show-InvPdfText -Path .\invoice.pdf        # see the page as the parser sees it

 HOW
   The engine is lib\InvPdfExtract.cs, compiled at run time by Add-Type using
   the compiler already inside the .NET Framework. Nothing is downloaded and
   nothing is installed - the same approach as NeXlsx.ps1 in the address repo.
================================================================================
#>

Set-StrictMode -Version 2.0

function Initialize-InvPdfEngine {
    <#
      Compiles the extractor once per session. Safe to call repeatedly.
    #>
    [CmdletBinding()]
    param([switch]$Force)

    if (-not $Force -and ('InvParse.PdfText' -as [type])) { return }

    $root = $PSScriptRoot
    if ([string]::IsNullOrEmpty($root)) { $root = (Get-Location).Path }
    $cs = Join-Path $root 'lib\InvPdfExtract.cs'
    if (-not (Test-Path -LiteralPath $cs)) {
        $cs = Join-Path $root 'lib/InvPdfExtract.cs'
    }
    if (-not (Test-Path -LiteralPath $cs)) {
        throw ("The PDF engine source is missing: lib\InvPdfExtract.cs`n" +
               "It must sit next to InvPdfText.ps1 - copy the whole folder, not just the .ps1 files.")
    }

    # -Encoding UTF8 matters: Windows PowerShell 5.1 reads files as the ANSI
    # code page by default, which silently corrupts any byte above 127.
    $src = Get-Content -LiteralPath $cs -Raw -Encoding UTF8
    try {
        Add-Type -TypeDefinition $src -Language CSharp -ErrorAction Stop
    }
    catch {
        throw ("Could not compile the PDF engine.`n" +
               "Windows PowerShell compiles it with the .NET Framework compiler, which needs a`n" +
               "writable TEMP folder. If your machine blocks that, run PowerShell from a folder`n" +
               "you own and try again.`n`nUnderlying error: " + $_.Exception.Message)
    }
}

<#
 Reads one PDF and returns every word with its position.

 Returns an object with:
   Ok        - $true when text came out
   Words     - InvParse.PdfWord[]  (Page, Text, X, Y, W, H, Size, Bold)
   Pages     - page count
   PageInfo  - per-page Width/Height in points
   Note      - why it is empty, when it is
   Scanned   - $true when the PDF has pages but no text at all (needs OCR)
   Encrypted - $true when the file carries /Encrypt
#>
function Get-InvPdfWords {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string] $Path
    )

    process {
        Initialize-InvPdfEngine

        $full = (Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue)
        if (-not $full) {
            return [pscustomobject]@{
                Ok = $false; Words = @(); Pages = 0; PageInfo = @()
                Note = 'File not found'; Scanned = $false; Encrypted = $false
                Path = $Path
            }
        }

        $res = $null
        try   { $res = [InvParse.PdfText]::Extract($full.ProviderPath) }
        catch {
            return [pscustomobject]@{
                Ok = $false; Words = @(); Pages = 0; PageInfo = @()
                Note = ('Unreadable PDF: ' + $_.Exception.Message)
                Scanned = $false; Encrypted = $false; Path = $full.ProviderPath
            }
        }

        $words = $res.Words
        $pages = $res.PageInfo.Count
        $note  = ''
        $scanned = $false

        if ($res.Error) { $note = $res.Error }
        elseif ($words.Count -eq 0) {
            if ($pages -gt 0) {
                $scanned = $true
                $note = 'No text layer - this is a scan or an image. Needs OCR.'
            }
            else { $note = 'No pages found - the file may be damaged or not a PDF.' }
        }
        if ($res.Encrypted -and $words.Count -eq 0 -and -not $scanned) {
            $note = 'The PDF is encrypted and its text could not be read.'
        }

        [pscustomobject]@{
            Ok        = ($words.Count -gt 0)
            Words     = $words
            Pages     = $pages
            PageInfo  = $res.PageInfo
            Note      = $note
            Scanned   = $scanned
            Encrypted = $res.Encrypted
            Path      = $full.ProviderPath
        }
    }
}

<#
 Prints the page the way the parser sees it: words placed back onto a character
 grid. This is the tool for "why did it not find the total?" - if the number is
 not here, no rule can find it.
#>
function Show-InvPdfText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [int] $Page = 0,                 # 0 = all pages
        [double] $CharWidth = 4.8,       # points per output column
        [double] $LineHeight = 9.5       # points per output row
    )

    $r = Get-InvPdfWords -Path $Path
    if (-not $r.Ok) {
        Write-Host ("  (nothing to show: " + $r.Note + ")") -ForegroundColor Yellow
        return
    }

    $pageNums = if ($Page -gt 0) { @($Page) } else { 1..$r.Pages }

    foreach ($pn in $pageNums) {
        $pw = $r.Words | Where-Object { $_.Page -eq $pn }
        if (-not $pw) { continue }

        $info = $r.PageInfo | Where-Object { $_.Number -eq $pn } | Select-Object -First 1
        $cols = [int][Math]::Ceiling($info.Width / $CharWidth) + 2
        $rows = [int][Math]::Ceiling($info.Height / $LineHeight) + 2

        Write-Host ""
        Write-Host ("  --- page {0} of {1}  ({2:N0} x {3:N0} pt, {4} words) ---" -f `
                    $pn, $r.Pages, $info.Width, $info.Height, $pw.Count) -ForegroundColor DarkCyan

        $grid = New-Object 'string[]' $rows
        for ($i = 0; $i -lt $rows; $i++) { $grid[$i] = '' }

        foreach ($w in $pw) {
            $row = [int][Math]::Round($w.Y / $LineHeight)
            if ($row -lt 0) { $row = 0 }
            if ($row -ge $rows) { $row = $rows - 1 }
            $col = [int][Math]::Round($w.X / $CharWidth)
            if ($col -lt 0) { $col = 0 }
            if ($col -gt $cols) { $col = $cols }

            $line = $grid[$row]
            if ($line.Length -lt $col) { $line = $line.PadRight($col) }
            # Never let two words fuse just because the grid is narrower than the
            # page - a missing space here reads as an extraction bug when it is not.
            elseif ($line.Length -ge $col -and $line.Length -gt 0) { $line = $line + ' ' }
            $grid[$row] = $line + $w.Text
        }

        foreach ($line in $grid) {
            if ($line.Trim().Length -gt 0) { Write-Host ('  ' + $line.TrimEnd()) }
        }
    }
}
