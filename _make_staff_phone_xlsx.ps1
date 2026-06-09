# Build a name + simulated-phone xlsx from รายละเอียดเจ้าหน้าและลูกจ้างสหกรณ์.xlsx
# Output columns: ชื่อ - สกุล | เบอร์โทร (sequential 10-digit text, 0000000001+)
$ErrorActionPreference = 'Stop'
Set-Location 'd:\POS'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$src = 'd:\POS\รายละเอียดเจ้าหน้าและลูกจ้างสหกรณ์.xlsx'
$out = 'd:\POS\รายชื่อเจ้าหน้าที่และลูกจ้าง-เบอร์จำลอง.xlsx'

# ---- read source (shared read so it works while open in Excel) ----
$fs = [System.IO.File]::Open($src, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
function ReadEntry($z, $name) {
  $e = $z.GetEntry($name)
  $r = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
  $t = $r.ReadToEnd(); $r.Dispose(); return $t
}
[xml]$ss = ReadEntry $zip 'xl/sharedStrings.xml'
$strings = New-Object System.Collections.Generic.List[string]
foreach ($si in $ss.sst.si) {
  if ($si.t -is [string]) { $strings.Add($si.t) }
  elseif ($si.t.'#text') { $strings.Add([string]$si.t.'#text') }
  else { $txt = ''; foreach ($r in $si.r) { $txt += [string]$r.t.'#text' }; $strings.Add($txt) }
}
[xml]$sh = ReadEntry $zip 'xl/worksheets/sheet1.xml'
$rows = $sh.worksheet.sheetData.row
$names = New-Object System.Collections.Generic.List[string]
foreach ($row in $rows) {
  if ([int]$row.r -lt 4) { continue }   # skip title rows 1-2 and header row 3
  $nameVal = $null
  foreach ($c in $row.c) {
    if ($c.r -match '^B\d+$') {
      if ($c.t -eq 's') { $nameVal = $strings[[int]$c.v] }
      elseif ($c.is) { $nameVal = [string]$c.is.t.'#text' }
      else { $nameVal = [string]$c.v }
    }
  }
  if ($nameVal) {
    $nameVal = ($nameVal -replace '\s+', ' ').Trim()
    if ($nameVal) { $names.Add($nameVal) }
  }
}
$zip.Dispose(); $fs.Dispose()
Write-Host "Extracted names: $($names.Count)"

# ---- build worksheet XML ----
function Esc($s) { ($s -replace '&', '&amp;') -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' }
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append('<row r="1"><c r="A1" t="inlineStr" s="1"><is><t xml:space="preserve">' + (Esc 'ชื่อ - สกุล') + '</t></is></c><c r="B1" t="inlineStr" s="1"><is><t xml:space="preserve">' + (Esc 'เบอร์โทร') + '</t></is></c></row>')
for ($i = 0; $i -lt $names.Count; $i++) {
  $r = $i + 2
  $phone = ([string]($i + 1)).PadLeft(10, '0')
  [void]$sb.Append('<row r="' + $r + '"><c r="A' + $r + '" t="inlineStr"><is><t xml:space="preserve">' + (Esc $names[$i]) + '</t></is></c><c r="B' + $r + '" t="inlineStr"><is><t xml:space="preserve">' + $phone + '</t></is></c></row>')
}
$sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="32" customWidth="1"/><col min="2" max="2" width="16" customWidth="1"/></cols><sheetData>' + $sb.ToString() + '</sheetData></worksheet>'

$stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="TH Sarabun New"/></font><font><b/><sz val="11"/><name val="TH Sarabun New"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>'

$workbookXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="เจ้าหน้าที่" sheetId="1" r:id="rId1"/></sheets></workbook>'

$workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'

$rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

$contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'

# ---- write output xlsx ----
if (Test-Path $out) { [System.IO.File]::Delete($out) }
$utf8 = New-Object System.Text.UTF8Encoding($false)
$ofs = [System.IO.File]::Open($out, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
$ozip = New-Object System.IO.Compression.ZipArchive($ofs, [System.IO.Compression.ZipArchiveMode]::Create)
function AddEntry($z, $name, $content, $enc) {
  $e = $z.CreateEntry($name)
  $w = New-Object System.IO.StreamWriter($e.Open(), $enc)
  $w.Write($content); $w.Dispose()
}
AddEntry $ozip '[Content_Types].xml' $contentTypes $utf8
AddEntry $ozip '_rels/.rels' $rootRels $utf8
AddEntry $ozip 'xl/workbook.xml' $workbookXml $utf8
AddEntry $ozip 'xl/_rels/workbook.xml.rels' $workbookRels $utf8
AddEntry $ozip 'xl/worksheets/sheet1.xml' $sheetXml $utf8
AddEntry $ozip 'xl/styles.xml' $stylesXml $utf8
$ozip.Dispose(); $ofs.Dispose()

Write-Host "Wrote: $out"
Write-Host ("First 3: " + (($names[0..2] | ForEach-Object -Begin { $n = 0 } -Process { $n++; "$($_)|$(([string]$n).PadLeft(10,'0'))" }) -join '  ::  '))
$last = $names.Count
Write-Host ("Last:    " + $names[$last-1] + "|" + ([string]$last).PadLeft(10,'0'))
