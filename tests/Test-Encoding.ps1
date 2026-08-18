Start-InvSuite 'Encoding - the files must load on Windows PowerShell 5.1'

<#
  Windows PowerShell 5.1 reads .ps1 files as the machine's ANSI code page, NOT
  as UTF-8, unless the file carries a UTF-8 BOM. Any byte above 127 is then
  read as the wrong character and the script fails to PARSE - which takes the
  whole tool down before a single line runs.

  This one bit for real: a pound sign and a euro sign inside a currency regex.

  The rule is therefore simple and absolute: every shipped file is pure ASCII.
  Characters above 127 are written as escapes - a backslash-u code in a
  .NET regex or a C# string literal - which are themselves plain ASCII.
#>

$shipped = @()
$shipped += @(Get-ChildItem -LiteralPath $script:InvRoot -Filter '*.ps1' -File)
$shipped += @(Get-ChildItem -LiteralPath (Join-Path $script:InvRoot 'tests') -Filter '*.ps1' -File)
$shipped += @(Get-ChildItem -LiteralPath (Join-Path $script:InvRoot 'lib') -Filter '*.cs' -File)

Assert-InvTrue 'there are files to check' ($shipped.Count -ge 12)

foreach ($f in $shipped) {
    $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
    $bad = 0
    $firstLine = 0
    $line = 1
    foreach ($b in $bytes) {
        if ($b -eq 10) { $line++ }
        if ($b -gt 127) {
            $bad++
            if ($firstLine -eq 0) { $firstLine = $line }
        }
    }
    if ($bad -gt 0) {
        Assert-InvEqual ($f.Name + ' is pure ASCII (first offending line ' + $firstLine + ')') 0 $bad
    }
    else {
        Assert-InvEqual ($f.Name + ' is pure ASCII') 0 $bad
    }
}

# and the escapes actually still do their job
Assert-InvEqual 'a pound amount still parses'  1234.56 (Convert-InvMoney ([char]0x00A3 + '1,234.56'))
Assert-InvEqual 'a euro amount still parses'   1234.56 (Convert-InvMoney ([char]0x20AC + '1,234.56'))
Assert-InvEqual 'a dollar amount still parses' 1234.56 (Convert-InvMoney '$1,234.56')
