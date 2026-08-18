<#
 Runs every suite.

     .\tests\Run-AllTests.ps1

 Needs nothing installed. Exits non-zero if anything failed, so it can be
 wired into a check before you trust a run.
#>
[CmdletBinding()]
param([string[]]$Only)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
if ([string]::IsNullOrEmpty($here)) { $here = (Get-Location).Path }
$script:InvRoot = Split-Path -Parent $here

foreach ($m in @('InvPdfText','InvDicts','InvLayout','InvValues','InvIdentity',
                 'InvFields','InvRules','InvRecord','InvColumns','InvXlsx')) {
    . (Join-Path $script:InvRoot ($m + '.ps1'))
}
. (Join-Path $here 'InvTest.ps1')

# The identity used by the tests. Written from the example if there is none, so
# the suite runs on a clean checkout without touching anybody's real profile.
$cfg = Join-Path $script:InvRoot 'config/MyCompany.json'
if (-not (Test-Path -LiteralPath $cfg)) {
    Copy-Item (Join-Path $script:InvRoot 'config/MyCompany.example.json') $cfg
    Write-Host '  (created config/MyCompany.json from the example for this test run)' -ForegroundColor DarkGray
}

Write-Host ''
Write-Host '  PDF INVOICE PARSER - TESTS' -ForegroundColor Cyan
Write-Host '  ---------------------------------------------------------------------------' -ForegroundColor DarkGray

$suites = @('Test-Values','Test-Layout','Test-Identity','Test-Rules','Test-EndToEnd','Test-Edge','Test-Xlsx')
if ($Only) { $suites = @($suites | Where-Object { $Only -contains $_ }) }

foreach ($s in $suites) {
    $p = Join-Path $here ($s + '.ps1')
    if (-not (Test-Path -LiteralPath $p)) { continue }
    . $p
}

$failed = Write-InvTestSummary
if ($failed -gt 0) { exit 1 }
exit 0
