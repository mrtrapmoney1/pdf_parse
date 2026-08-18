<#
 A tiny assertion framework. No Pester, no modules to install - the test suite
 has to run on a locked-down work machine with nothing but PowerShell.
#>

Set-StrictMode -Version 2.0

$script:InvPass = 0
$script:InvFail = 0
$script:InvFailures = [System.Collections.ArrayList]::new()
$script:InvSuite = ''

function Start-InvSuite {
    param([string]$Name)
    $script:InvSuite = $Name
    Write-Host ''
    Write-Host ('  ' + $Name) -ForegroundColor Cyan
}

function Assert-InvEqual {
    param([string]$What, $Expected, $Actual)

    $ok = $false
    if ($null -eq $Expected -and $null -eq $Actual) { $ok = $true }
    elseif ($null -eq $Expected -or $null -eq $Actual) { $ok = $false }
    elseif ($Expected -is [double] -or $Actual -is [double]) {
        $ok = ([Math]::Abs([double]$Expected - [double]$Actual) -lt 0.005)
    }
    elseif ($Expected -is [datetime]) { $ok = ([datetime]$Expected -eq [datetime]$Actual) }
    else { $ok = ([string]$Expected -eq [string]$Actual) }

    if ($ok) {
        $script:InvPass++
    }
    else {
        $script:InvFail++
        $msg = ('{0} / {1}: expected <{2}> but got <{3}>' -f $script:InvSuite, $What, $Expected, $Actual)
        [void]$script:InvFailures.Add($msg)
        Write-Host ('    FAIL  ' + $What + '  expected <' + $Expected + '> got <' + $Actual + '>') -ForegroundColor Red
    }
}

function Assert-InvTrue {
    param([string]$What, $Condition)
    if ($Condition) { $script:InvPass++ }
    else {
        $script:InvFail++
        $msg = ('{0} / {1}: expected true' -f $script:InvSuite, $What)
        [void]$script:InvFailures.Add($msg)
        Write-Host ('    FAIL  ' + $What) -ForegroundColor Red
    }
}

function Assert-InvNull {
    param([string]$What, $Value)
    if ($null -eq $Value -or ([string]$Value) -eq '') { $script:InvPass++ }
    else {
        $script:InvFail++
        $msg = ('{0} / {1}: expected nothing, got <{2}>' -f $script:InvSuite, $What, $Value)
        [void]$script:InvFailures.Add($msg)
        Write-Host ('    FAIL  ' + $What + '  expected nothing, got <' + $Value + '>') -ForegroundColor Red
    }
}

function Write-InvTestSummary {
    Write-Host ''
    Write-Host '  ---------------------------------------------------------------------------' -ForegroundColor DarkGray
    if ($script:InvFail -eq 0) {
        Write-Host ('  ALL PASS   {0} assertions' -f $script:InvPass) -ForegroundColor Green
    }
    else {
        Write-Host ('  {0} PASSED   {1} FAILED' -f $script:InvPass, $script:InvFail) -ForegroundColor Red
        Write-Host ''
        foreach ($f in $script:InvFailures) { Write-Host ('   . ' + $f) -ForegroundColor Red }
    }
    Write-Host ''
    return $script:InvFail
}
