<#
================================================================================
 InvXlsx.ps1
 Writes a real .xlsx. No Excel, no ImportExcel module, no COM, no install.

 HOW
   An .xlsx IS a zip of XML parts. We build the parts as text and zip them with
   System.IO.Compression, which is part of .NET. That means the tool produces a
   workbook on a machine with no Office installed at all, and it never has to
   drive Excel over COM - which is slow, fragile, and leaves stray EXCEL.EXE
   processes behind when anything goes wrong.

 WHAT YOU GET
   - a frozen, filtered header row
   - money, date and percentage cells formatted as numbers, not as text, so
     they total and sort correctly in the workpaper
   - amber shading on any cell the parser was not confident about
   - the Source cell hyperlinked to the PDF it came from
   - a Log sheet recording what happened to every file in the run
================================================================================
#>

Set-StrictMode -Version 2.0

function Initialize-InvZip {
    if (-not ('System.IO.Compression.ZipArchive' -as [type])) {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    }
    if (-not ('System.IO.Compression.ZipFile' -as [type])) {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    }
}

function ConvertTo-InvXmlText {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $s = $Text -replace '&', '&amp;'
    $s = $s -replace '<', '&lt;'
    $s = $s -replace '>', '&gt;'
    $s = $s -replace '"', '&quot;'
    # control characters are illegal in XML and will make Excel refuse the file
    $s = [regex]::Replace($s, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    return $s
}

function ConvertTo-InvColumnLetter {
    param([int]$Index)          # 1-based
    $s = ''
    $n = $Index
    while ($n -gt 0) {
        $r = ($n - 1) % 26
        $s = [char](65 + $r) + $s
        $n = [int](($n - $r) / 26)
    }
    return $s
}

# style ids, matching the cellXfs list built in Get-InvStylesXml
$script:InvStyle = @{
    General = 0; Header = 1; Text = 2; Money = 3; Date = 4; Pct = 5; Int = 6
    TextLow = 7; MoneyLow = 8; DateLow = 9; PctLow = 10; IntLow = 11
    Note    = 12
}

function Get-InvStyleId {
    param([string]$Type, [bool]$Low, $Map = $null)
    if ($null -eq $Map) { $Map = $script:InvStyle }
    switch ($Type) {
        'money' { if ($Low) { return $Map.MoneyLow } else { return $Map.Money } }
        'date'  { if ($Low) { return $Map.DateLow }  else { return $Map.Date } }
        'pct'   { if ($Low) { return $Map.PctLow }   else { return $Map.Pct } }
        'int'   { if ($Low) { return $Map.IntLow }   else { return $Map.Int } }
        default { if ($Low) { return $Map.TextLow }  else { return $Map.Text } }
    }
}

function Get-InvStylesXml {
@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="3">
<numFmt numFmtId="164" formatCode="#,##0.00"/>
<numFmt numFmtId="165" formatCode="mm/dd/yyyy"/>
<numFmt numFmtId="166" formatCode="0.000%"/>
</numFmts>
<fonts count="4">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
<font><sz val="11"/><color rgb="FF7F6000"/><name val="Calibri"/><family val="2"/></font>
<font><i/><sz val="10"/><color rgb="FF595959"/><name val="Calibri"/><family val="2"/></font>
</fonts>
<fills count="4">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
<border><left/><right/><top/><bottom/><diagonal/></border>
<border><left/><right/><top/><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="13">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="165" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="166" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="3" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="0"/>
</styleSheet>
'@
}

<#
 Builds one worksheet.

 $Columns  : objects with Key, Header, Type ('text'|'money'|'date'|'pct'|'int'), Width
 $Rows     : hashtables mapping Key -> @{ V = <value>; Low = <bool> }
             plus an optional '_Link' entry giving a file path to hyperlink the
             Source cell to.
#>
function Get-InvSheetXml {
    param($Columns, $Rows, [string]$LinkColumnKey = 'Source', $StyleMap = $null)

    # When a sheet is appended to somebody else's workbook, our style ids are
    # not 0..12 any more - they land after whatever styles that workbook
    # already had. The caller passes the remapped ids in.
    if ($null -eq $StyleMap) { $StyleMap = $script:InvStyle }

    $cols = @($Columns)
    $rows = @($Rows)
    $lastCol = ConvertTo-InvColumnLetter $cols.Count
    $lastRow = $rows.Count + 1

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sb.Append('<dimension ref="A1:' + $lastCol + $lastRow + '"/>')
    [void]$sb.Append('<sheetViews><sheetView tabSelected="1" workbookViewId="0">')
    [void]$sb.Append('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>')
    [void]$sb.Append('<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>')
    [void]$sb.Append('</sheetView></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/>')

    [void]$sb.Append('<cols>')
    for ($i = 0; $i -lt $cols.Count; $i++) {
        $w = 14
        if ($cols[$i].PSObject.Properties.Name -contains 'Width' -and $cols[$i].Width) { $w = $cols[$i].Width }
        [void]$sb.Append('<col min="' + ($i + 1) + '" max="' + ($i + 1) + '" width="' + $w + '" customWidth="1"/>')
    }
    [void]$sb.Append('</cols>')

    [void]$sb.Append('<sheetData>')

    # header
    [void]$sb.Append('<row r="1" ht="30" customHeight="1">')
    for ($i = 0; $i -lt $cols.Count; $i++) {
        $ref = (ConvertTo-InvColumnLetter ($i + 1)) + '1'
        [void]$sb.Append('<c r="' + $ref + '" s="' + $StyleMap.Header + '" t="inlineStr"><is><t xml:space="preserve">' +
                         (ConvertTo-InvXmlText $cols[$i].Header) + '</t></is></c>')
    }
    [void]$sb.Append('</row>')

    $links = [System.Collections.ArrayList]::new()

    for ($r = 0; $r -lt $rows.Count; $r++) {
        $rowNum = $r + 2
        [void]$sb.Append('<row r="' + $rowNum + '">')
        $row = $rows[$r]

        for ($i = 0; $i -lt $cols.Count; $i++) {
            $key = $cols[$i].Key
            $type = $cols[$i].Type
            $ref = (ConvertTo-InvColumnLetter ($i + 1)) + $rowNum

            $v = $null; $low = $false
            if ($row.ContainsKey($key)) {
                $cell = $row[$key]
                $v = $cell.V
                if ($cell.ContainsKey('Low')) { $low = [bool]$cell.Low }
            }

            $sid = Get-InvStyleId -Type $type -Low $low -Map $StyleMap

            if ($null -eq $v -or ([string]$v) -eq '') {
                [void]$sb.Append('<c r="' + $ref + '" s="' + $sid + '"/>')
                continue
            }

            switch ($type) {
                'money' {
                    [void]$sb.Append('<c r="' + $ref + '" s="' + $sid + '"><v>' +
                        ([double]$v).ToString('0.####', [System.Globalization.CultureInfo]::InvariantCulture) + '</v></c>')
                }
                'int' {
                    [void]$sb.Append('<c r="' + $ref + '" s="' + $sid + '"><v>' + ([int]$v) + '</v></c>')
                }
                'pct' {
                    # stored as a fraction so Excel's percent format shows 7.000%
                    $frac = ([double]$v) / 100.0
                    [void]$sb.Append('<c r="' + $ref + '" s="' + $sid + '"><v>' +
                        $frac.ToString('0.########', [System.Globalization.CultureInfo]::InvariantCulture) + '</v></c>')
                }
                'date' {
                    $dt = [datetime]$v
                    $serial = ($dt - [datetime]'1899-12-30').Days
                    [void]$sb.Append('<c r="' + $ref + '" s="' + $sid + '"><v>' + $serial + '</v></c>')
                }
                default {
                    [void]$sb.Append('<c r="' + $ref + '" s="' + $sid + '" t="inlineStr"><is><t xml:space="preserve">' +
                        (ConvertTo-InvXmlText ([string]$v)) + '</t></is></c>')
                }
            }

            if ($key -eq $LinkColumnKey -and $row.ContainsKey('_Link') -and $row['_Link']) {
                [void]$links.Add([pscustomobject]@{ Ref = $ref; Target = [string]$row['_Link'] })
            }
        }
        [void]$sb.Append('</row>')
    }

    [void]$sb.Append('</sheetData>')
    [void]$sb.Append('<autoFilter ref="A1:' + $lastCol + $lastRow + '"/>')

    if ($links.Count -gt 0) {
        [void]$sb.Append('<hyperlinks>')
        for ($i = 0; $i -lt $links.Count; $i++) {
            [void]$sb.Append('<hyperlink ref="' + $links[$i].Ref + '" r:id="rId' + ($i + 1) + '"/>')
        }
        [void]$sb.Append('</hyperlinks>')
    }

    [void]$sb.Append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    [void]$sb.Append('</worksheet>')

    return [pscustomobject]@{ Xml = $sb.ToString(); Links = @($links) }
}

function Get-InvLogSheetXml {
    param($LogLines, $StyleMap = $null)
    if ($null -eq $StyleMap) { $StyleMap = $script:InvStyle }

    $lines = @($LogLines)
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/>')
    [void]$sb.Append('<cols><col min="1" max="1" width="34" customWidth="1"/><col min="2" max="2" width="120" customWidth="1"/></cols>')
    [void]$sb.Append('<sheetData>')
    [void]$sb.Append('<row r="1"><c r="A1" s="' + $StyleMap.Header + '" t="inlineStr"><is><t>File</t></is></c>' +
                     '<c r="B1" s="' + $StyleMap.Header + '" t="inlineStr"><is><t>What happened</t></is></c></row>')
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $r = $i + 2
        [void]$sb.Append('<row r="' + $r + '">')
        [void]$sb.Append('<c r="A' + $r + '" s="' + $StyleMap.Text + '" t="inlineStr"><is><t xml:space="preserve">' +
                         (ConvertTo-InvXmlText ([string]$lines[$i].File)) + '</t></is></c>')
        [void]$sb.Append('<c r="B' + $r + '" s="' + $StyleMap.Text + '" t="inlineStr"><is><t xml:space="preserve">' +
                         (ConvertTo-InvXmlText ([string]$lines[$i].Message)) + '</t></is></c>')
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    [void]$sb.Append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

<#
 Writes the workbook. Overwrites $Path.
#>
function Export-InvWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] $Columns,
        [Parameter(Mandatory)] [AllowEmptyCollection()] $Rows,
        $LogLines = @(),
        [string] $SheetName = 'Invoices'
    )

    Initialize-InvZip

    $sheet = Get-InvSheetXml -Columns $Columns -Rows $Rows
    $logXml = Get-InvLogSheetXml -LogLines $LogLines

    $full = $Path
    if (-not [System.IO.Path]::IsPathRooted($full)) {
        $full = Join-Path (Get-Location).Path $Path
    }
    $dir = Split-Path -Parent $full
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    if (Test-Path -LiteralPath $full) { Remove-Item -LiteralPath $full -Force }

    $parts = [ordered]@{}

    $parts['[Content_Types].xml'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'@

    $parts['_rels/.rels'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

    $parts['xl/workbook.xml'] =
'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
'<sheets>' +
'<sheet name="' + (ConvertTo-InvXmlText $SheetName) + '" sheetId="1" r:id="rId1"/>' +
'<sheet name="Log" sheetId="2" r:id="rId2"/>' +
'</sheets></workbook>'

    $parts['xl/_rels/workbook.xml.rels'] = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@

    $parts['xl/styles.xml'] = (Get-InvStylesXml)
    $parts['xl/worksheets/sheet1.xml'] = $sheet.Xml
    $parts['xl/worksheets/sheet2.xml'] = $logXml

    if ($sheet.Links.Count -gt 0) {
        $rels = [System.Text.StringBuilder]::new()
        [void]$rels.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$rels.Append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
        for ($i = 0; $i -lt $sheet.Links.Count; $i++) {
            $target = ConvertTo-InvXmlText ($sheet.Links[$i].Target -replace '\\', '/')
            [void]$rels.Append('<Relationship Id="rId' + ($i + 1) +
                '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///' +
                $target + '" TargetMode="External"/>')
        }
        [void]$rels.Append('</Relationships>')
        $parts['xl/worksheets/_rels/sheet1.xml.rels'] = $rels.ToString()
    }

    $fs = $null; $zip = $null
    try {
        $fs = [System.IO.File]::Open($full, [System.IO.FileMode]::CreateNew)
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        $utf8 = [System.Text.UTF8Encoding]::new($false)

        foreach ($name in $parts.Keys) {
            $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $es = $entry.Open()
            try {
                $sw = [System.IO.StreamWriter]::new($es, $utf8)
                try { $sw.Write([string]$parts[$name]) } finally { $sw.Dispose() }
            }
            finally { $es.Dispose() }
        }
    }
    finally {
        if ($zip) { $zip.Dispose() }
        if ($fs)  { $fs.Dispose() }
    }

    return $full
}

<#
 The CSV safety net. Written on every run, so a workbook problem can never cost
 you the extraction.
#>
function Export-InvCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] $Columns,
        [Parameter(Mandatory)] [AllowEmptyCollection()] $Rows
    )

    $out = [System.Collections.ArrayList]::new()
    foreach ($row in @($Rows)) {
        $o = [ordered]@{}
        foreach ($c in @($Columns)) {
            $v = ''
            if ($row.ContainsKey($c.Key)) {
                $raw = $row[$c.Key].V
                if ($null -ne $raw) {
                    if ($raw -is [datetime]) { $v = $raw.ToString('MM/dd/yyyy') }
                    elseif ($c.Type -eq 'money') { $v = ([double]$raw).ToString('0.00', [System.Globalization.CultureInfo]::InvariantCulture) }
                    else { $v = [string]$raw }
                }
            }
            $o[$c.Header] = $v
        }
        [void]$out.Add([pscustomobject]$o)
    }

    if ($out.Count -eq 0) { return $null }
    $out | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    return $Path
}

# ------------------------------------------------- appending to a workbook you
#                                                    already have

<#
 Adds a sheet to an EXISTING .xlsx without disturbing anything already in it.

 The awkward part is styles. Cell formatting is referenced by index into the
 workbook's own style table, so our "money" style is not index 3 in somebody
 else's file. This merges our fonts, fills, borders and formats onto the end of
 the host workbook's style table and remaps our ids to wherever they landed.

 The original file is copied to <name>.bak.xlsx before anything is written.
#>
function Add-InvWorksheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] $Columns,
        [Parameter(Mandatory)] [AllowEmptyCollection()] $Rows,
        [string] $SheetName = 'Invoices',
        $LogLines = @()
    )

    Initialize-InvZip

    $full = (Resolve-Path -LiteralPath $Path).ProviderPath
    if ([System.IO.Path]::GetExtension($full).ToLowerInvariant() -ne '.xlsx') {
        throw ("Only .xlsx workbooks can be added to. This is " + [System.IO.Path]::GetExtension($full) +
               " - save it as .xlsx first, or write to a new workbook instead.")
    }

    # never modify the only copy
    $bak = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($full),
           [System.IO.Path]::GetFileNameWithoutExtension($full) + '.bak.xlsx')
    Copy-Item -LiteralPath $full -Destination $bak -Force

    # read every part into memory
    $parts = [ordered]@{}
    $zin = [System.IO.Compression.ZipFile]::OpenRead($full)
    try {
        foreach ($e in $zin.Entries) {
            if ([string]::IsNullOrEmpty($e.Name)) { continue }
            $st = $e.Open()
            try {
                $ms = [System.IO.MemoryStream]::new()
                $st.CopyTo($ms)
                $parts[$e.FullName] = $ms.ToArray()
            }
            finally { $st.Dispose() }
        }
    }
    finally { $zin.Dispose() }

    $utf8 = [System.Text.UTF8Encoding]::new($false)
    $textOf = { param($name) if ($parts.Contains($name)) { [System.Text.Encoding]::UTF8.GetString($parts[$name]) } else { '' } }

    $wbXml   = & $textOf 'xl/workbook.xml'
    $relsXml = & $textOf 'xl/_rels/workbook.xml.rels'
    $ctXml   = & $textOf '[Content_Types].xml'
    if (-not $wbXml -or -not $relsXml -or -not $ctXml) {
        throw 'That file does not look like a workbook this tool can add to.'
    }

    # a free sheet part name, rId and sheetId
    $n = 1
    while ($parts.Contains('xl/worksheets/sheet' + $n + '.xml')) { $n++ }
    $sheetPart = 'xl/worksheets/sheet' + $n + '.xml'

    $maxRid = 0
    foreach ($m in [regex]::Matches($relsXml, 'Id="rId(\d+)"')) {
        $v = [int]$m.Groups[1].Value
        if ($v -gt $maxRid) { $maxRid = $v }
    }
    $rid = 'rId' + ($maxRid + 1)

    $maxSheetId = 0
    foreach ($m in [regex]::Matches($wbXml, 'sheetId="(\d+)"')) {
        $v = [int]$m.Groups[1].Value
        if ($v -gt $maxSheetId) { $maxSheetId = $v }
    }

    # a sheet name Excel will accept, and one that is not already taken
    $safe = [regex]::Replace($SheetName, '[\\/*?:\[\]]', '-')
    if ($safe.Length -gt 31) { $safe = $safe.Substring(0, 31) }
    $existingNames = @()
    foreach ($m in [regex]::Matches($wbXml, '<sheet[^>]*name="([^"]*)"')) { $existingNames += $m.Groups[1].Value }
    $try = $safe; $k = 2
    while ($existingNames -contains $try) {
        $suffix = ' (' + $k + ')'
        $try = $safe.Substring(0, [Math]::Min($safe.Length, 31 - $suffix.Length)) + $suffix
        $k++
    }
    $safe = $try

    # merge our styles onto the end of the host workbook's style table
    $stylesXml = & $textOf 'xl/styles.xml'
    $merged = Merge-InvStyles -StylesXml $stylesXml
    $parts['xl/styles.xml'] = $utf8.GetBytes($merged.Xml)

    $sheet = Get-InvSheetXml -Columns $Columns -Rows $Rows -StyleMap $merged.Map
    $parts[$sheetPart] = $utf8.GetBytes($sheet.Xml)

    if ($sheet.Links.Count -gt 0) {
        $rb = [System.Text.StringBuilder]::new()
        [void]$rb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$rb.Append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
        for ($i = 0; $i -lt $sheet.Links.Count; $i++) {
            $target = ConvertTo-InvXmlText ($sheet.Links[$i].Target -replace '\\', '/')
            [void]$rb.Append('<Relationship Id="rId' + ($i + 1) +
                '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="file:///' +
                $target + '" TargetMode="External"/>')
        }
        [void]$rb.Append('</Relationships>')
        $parts['xl/worksheets/_rels/sheet' + $n + '.xml.rels'] = $utf8.GetBytes($rb.ToString())
    }

    # register the sheet
    $wbXml = $wbXml -replace '</sheets>', ('<sheet name="' + (ConvertTo-InvXmlText $safe) +
             '" sheetId="' + ($maxSheetId + 1) + '" r:id="' + $rid + '"/></sheets>')
    $parts['xl/workbook.xml'] = $utf8.GetBytes($wbXml)

    $relsXml = $relsXml -replace '</Relationships>', ('<Relationship Id="' + $rid +
               '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' +
               $n + '.xml"/></Relationships>')
    $parts['xl/_rels/workbook.xml.rels'] = $utf8.GetBytes($relsXml)

    $ctXml = $ctXml -replace '</Types>', ('<Override PartName="/' + $sheetPart +
             '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    $parts['[Content_Types].xml'] = $utf8.GetBytes($ctXml)

    # rewrite the package
    $tmp = $full + '.tmp'
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }

    $fs = $null; $zip = $null
    try {
        $fs = [System.IO.File]::Open($tmp, [System.IO.FileMode]::CreateNew)
        $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($name in $parts.Keys) {
            $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $es = $entry.Open()
            try { $es.Write($parts[$name], 0, $parts[$name].Length) } finally { $es.Dispose() }
        }
    }
    finally {
        if ($zip) { $zip.Dispose() }
        if ($fs)  { $fs.Dispose() }
    }

    Move-Item -LiteralPath $tmp -Destination $full -Force
    return $full
}


<#
 Appends child elements to a style collection, handling BOTH forms a workbook
 can use: a normal <fonts count="3">...</fonts> and a self-closing
 <numFmts count="0"/>. Miss the self-closing form and the new entries are
 silently dropped, leaving cells pointing at formats that do not exist.
#>
function Add-InvStyleElements {
    param(
        [string] $Xml,
        [string] $Tag,
        [string] $Children,
        [int]    $Added,
        [switch] $CreateIfMissing
    )

    $selfClosing = [regex]::Match($Xml, '<' + $Tag + '(\s[^>]*?)?/>')
    if ($selfClosing.Success) {
        $existing = 0
        $cm = [regex]::Match($selfClosing.Value, 'count="(\d+)"')
        if ($cm.Success) { $existing = [int]$cm.Groups[1].Value }
        $replacement = '<' + $Tag + ' count="' + ($existing + $Added) + '">' + $Children + '</' + $Tag + '>'
        return $Xml.Remove($selfClosing.Index, $selfClosing.Length).Insert($selfClosing.Index, $replacement)
    }

    if ($Xml -match ('<' + $Tag + '(\s[^>]*)?>')) {
        $x2 = [regex]::Replace($Xml, '<' + $Tag + '([^>]*)count="(\d+)"',
              { param($m) '<' + $Tag + $m.Groups[1].Value + 'count="' + ([int]$m.Groups[2].Value + $Added) + '"' }, 1)
        return [regex]::Replace($x2, '</' + $Tag + '>', ($Children + '</' + $Tag + '>'), 1)
    }

    if ($CreateIfMissing) {
        return [regex]::Replace($Xml, '(<styleSheet[^>]*>)',
               ('$1<' + $Tag + ' count="' + $Added + '">' + $Children + '</' + $Tag + '>'), 1)
    }
    return $Xml
}

<#
 Appends our formatting to an existing styles.xml and reports where each of our
 styles ended up.
#>
function Merge-InvStyles {
    param([string]$StylesXml)

    if ([string]::IsNullOrWhiteSpace($StylesXml)) {
        return [pscustomobject]@{ Xml = (Get-InvStylesXml); Map = $script:InvStyle }
    }

    $x = $StylesXml

    $countOf = {
        param($tag)
        $m = [regex]::Match($x, '<' + $tag + '[^>]*count="(\d+)"')
        if ($m.Success) { return [int]$m.Groups[1].Value }
        # count is optional in the schema; count the children instead
        $body = [regex]::Match($x, '<' + $tag + '\b[^>]*>(.*?)</' + $tag + '>', 'Singleline')
        if (-not $body.Success) { return 0 }
        $inner = $body.Groups[1].Value
        $singular = $tag.TrimEnd('s')
        if ($tag -eq 'cellXfs') { $singular = 'xf' }
        if ($tag -eq 'numFmts') { $singular = 'numFmt' }
        return @([regex]::Matches($inner, '<' + $singular + '\b')).Count
    }

    $fontBase   = & $countOf 'fonts'
    $fillBase   = & $countOf 'fills'
    $borderBase = & $countOf 'borders'
    $xfBase     = & $countOf 'cellXfs'

    # Custom number formats start at 164. They must be allocated CONTIGUOUSLY
    # from there: Excel treats numFmtId as an id, but several readers (openpyxl
    # among them) index the custom list by position, and a gap makes them throw
    # when opening the file. So take the next three after the highest custom id
    # already present.
    $maxCustom = 163
    foreach ($m in [regex]::Matches($x, 'numFmtId="(\d+)"')) {
        $v = [int]$m.Groups[1].Value
        if ($v -ge 164 -and $v -gt $maxCustom) { $maxCustom = $v }
    }
    $fmtMoney = $maxCustom + 1
    $fmtDate  = $maxCustom + 2
    $fmtPct   = $maxCustom + 3

    $newNumFmts =
        '<numFmt numFmtId="' + $fmtMoney + '" formatCode="#,##0.00"/>' +
        '<numFmt numFmtId="' + $fmtDate  + '" formatCode="mm/dd/yyyy"/>' +
        '<numFmt numFmtId="' + $fmtPct   + '" formatCode="0.000%"/>'

    $x = Add-InvStyleElements -Xml $x -Tag 'numFmts' -Children $newNumFmts -Added 3 -CreateIfMissing

    $newFonts =
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>' +
        '<font><sz val="11"/><color rgb="FF7F6000"/><name val="Calibri"/><family val="2"/></font>'
    $x = Add-InvStyleElements -Xml $x -Tag 'fonts' -Children $newFonts -Added 2 -CreateIfMissing

    $newFills =
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>'
    $x = Add-InvStyleElements -Xml $x -Tag 'fills' -Children $newFills -Added 2 -CreateIfMissing

    $newBorders = '<border><left/><right/><top/><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>'
    $x = Add-InvStyleElements -Xml $x -Tag 'borders' -Children $newBorders -Added 1 -CreateIfMissing

    $fHead = $fontBase       # bold white
    $fLow  = $fontBase + 1   # amber text
    $flHead = $fillBase      # header fill
    $flLow  = $fillBase + 1  # amber fill
    $bThin  = $borderBase

    $xf = @(
        ('<xf numFmtId="0" fontId="' + $fHead + '" fillId="' + $flHead + '" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>')
        ('<xf numFmtId="0" fontId="0" fillId="0" borderId="' + $bThin + '" xfId="0" applyBorder="1"/>')
        ('<xf numFmtId="' + $fmtMoney + '" fontId="0" fillId="0" borderId="' + $bThin + '" xfId="0" applyNumberFormat="1" applyBorder="1"/>')
        ('<xf numFmtId="' + $fmtDate  + '" fontId="0" fillId="0" borderId="' + $bThin + '" xfId="0" applyNumberFormat="1" applyBorder="1"/>')
        ('<xf numFmtId="' + $fmtPct   + '" fontId="0" fillId="0" borderId="' + $bThin + '" xfId="0" applyNumberFormat="1" applyBorder="1"/>')
        ('<xf numFmtId="3" fontId="0" fillId="0" borderId="' + $bThin + '" xfId="0" applyNumberFormat="1" applyBorder="1"/>')
        ('<xf numFmtId="0" fontId="' + $fLow + '" fillId="' + $flLow + '" borderId="' + $bThin + '" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>')
        ('<xf numFmtId="' + $fmtMoney + '" fontId="' + $fLow + '" fillId="' + $flLow + '" borderId="' + $bThin + '" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>')
        ('<xf numFmtId="' + $fmtDate  + '" fontId="' + $fLow + '" fillId="' + $flLow + '" borderId="' + $bThin + '" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>')
        ('<xf numFmtId="' + $fmtPct   + '" fontId="' + $fLow + '" fillId="' + $flLow + '" borderId="' + $bThin + '" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>')
        ('<xf numFmtId="3" fontId="' + $fLow + '" fillId="' + $flLow + '" borderId="' + $bThin + '" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>')
    )
    $x = Add-InvStyleElements -Xml $x -Tag 'cellXfs' -Children ($xf -join '') -Added 11 -CreateIfMissing

    $map = @{
        General  = 0
        Header   = $xfBase
        Text     = $xfBase + 1
        Money    = $xfBase + 2
        Date     = $xfBase + 3
        Pct      = $xfBase + 4
        Int      = $xfBase + 5
        TextLow  = $xfBase + 6
        MoneyLow = $xfBase + 7
        DateLow  = $xfBase + 8
        PctLow   = $xfBase + 9
        IntLow   = $xfBase + 10
        Note     = 0
    }

    return [pscustomobject]@{ Xml = $x; Map = $map }
}
