<#
Dump the facts of any Word template (.docx) so a profile can be written from evidence, not guesses.
Writes <OutDir>/template_probe.json (complete) and template_probe.md (compact, for reading).

  pwsh -File probe_template.ps1 -TemplatePath template.docx -OutDir probe

Facts only - no decisions. Deciding which paragraph is the "title" or the "body" sample is the
agent's job (references/new_journal_profile.md); validate_profile.ps1 then checks those decisions.
#>
param([Parameter(Mandatory)] [string]$TemplatePath, [Parameter(Mandatory)] [string]$OutDir)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$TemplatePath = (Resolve-Path $TemplatePath).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
$zip = [IO.Compression.ZipFile]::OpenRead($TemplatePath)
$parts = @{}
try { foreach ($e in $zip.Entries) { if ($e.FullName -match '\.xml$|\.rels$') { $r = [IO.StreamReader]::new($e.Open()); try { $parts[$e.FullName] = $r.ReadToEnd() } finally { $r.Dispose() } } } } finally { $zip.Dispose() }
function Load([string]$xml) { $d = [Xml.XmlDocument]::new(); $d.PreserveWhitespace = $true; $d.LoadXml($xml); $ns = [Xml.XmlNamespaceManager]::new($d.NameTable); $ns.AddNamespace('w', $W); $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math'); $ns.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'); return @($d, $ns) }
function A($node, [string]$xp, [string]$attr) { if (-not $node) { return $null }; $n = $node.SelectSingleNode($xp, $script:ns); if ($n) { return $n.GetAttribute($attr, $W) } }

$doc, $ns = Load $parts['word/document.xml']
$styles = @{}
if ($parts['word/styles.xml']) {
    $sd, $sns = Load $parts['word/styles.xml']
    foreach ($s in $sd.SelectNodes('//w:style', $sns)) {
        $id = $s.GetAttribute('styleId', $W); $sz = $s.SelectSingleNode('w:rPr/w:sz', $sns)
        $styles[$id] = [ordered]@{ id = $id; name = (A $s 'w:name' 'val'); type = $s.GetAttribute('type', $W); based_on = (A $s 'w:basedOn' 'val')
            size_pt = $(if ($sz) { [double]$sz.GetAttribute('val', $W) / 2 }); east_asia_font = (A $s 'w:rPr/w:rFonts' 'eastAsia'); latin_font = (A $s 'w:rPr/w:rFonts' 'ascii'); bold = [bool]$s.SelectSingleNode('w:rPr/w:b', $sns) }
    }
}

# sections
$sections = @(); $k = 0
foreach ($sp in $doc.SelectNodes('//w:sectPr', $ns)) {
    $pg = $sp.SelectSingleNode('w:pgSz', $ns); $mar = $sp.SelectSingleNode('w:pgMar', $ns); $cols = $sp.SelectSingleNode('w:cols', $ns)
    $owner = $sp.ParentNode; $loc = if ($owner.LocalName -eq 'pPr') { $i = 1 + @($owner.ParentNode.SelectNodes('preceding-sibling::w:p', $ns)).Count; "/w:document[1]/w:body[1]/w:p[$i]/w:pPr[1]/w:sectPr[1]" } else { '/w:document[1]/w:body[1]/w:sectPr[1]' }
    $sections += [ordered]@{ index = $k; locator = $loc; type = (A $sp 'w:type' 'val'); width_twips = [int]$pg.GetAttribute('w', $W); height_twips = [int]$pg.GetAttribute('h', $W)
        margins_twips = [ordered]@{ top = [int]$mar.GetAttribute('top', $W); right = [int]$mar.GetAttribute('right', $W); bottom = [int]$mar.GetAttribute('bottom', $W); left = [int]$mar.GetAttribute('left', $W); header = [int]$mar.GetAttribute('header', $W); footer = [int]$mar.GetAttribute('footer', $W) }
        columns = $(if ($cols -and $cols.GetAttribute('num', $W)) { [int]$cols.GetAttribute('num', $W) } else { 1 }); column_space_twips = $(if ($cols -and $cols.GetAttribute('space', $W)) { [int]$cols.GetAttribute('space', $W) })
        title_page = [bool]$sp.SelectSingleNode('w:titlePg', $ns); headers = @($sp.SelectNodes('w:headerReference', $ns) | ForEach-Object { "$($_.GetAttribute('type', $W))" }) }
    $k++
}

# body paragraphs (top level, in order)
$paras = @(); $i = 0
foreach ($p in $doc.SelectNodes('/w:document/w:body/w:p', $ns)) {
    $i++
    $text = (($p.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '')
    $runs = @($p.SelectNodes('w:r[w:t]', $ns))
    $sizes = @($runs | ForEach-Object { $s = $_.SelectSingleNode('w:rPr/w:sz', $ns); if ($s) { [double]$s.GetAttribute('val', $W) / 2 } } | Sort-Object -Unique)
    $ea = @($runs | ForEach-Object { A $_ 'w:rPr/w:rFonts' 'eastAsia' } | Where-Object { $_ } | Sort-Object -Unique)
    $lat = @($runs | ForEach-Object { A $_ 'w:rPr/w:rFonts' 'ascii' } | Where-Object { $_ } | Sort-Object -Unique)
    $styleId = A $p 'w:pPr/w:pStyle' 'val'
    $paras += [ordered]@{ index = $i; locator = "/w:document[1]/w:body[1]/w:p[$i]"; style_id = $styleId; style_name = $(if ($styleId -and $styles[$styleId]) { $styles[$styleId].name })
        style_size_pt = $(if ($styleId -and $styles[$styleId]) { $styles[$styleId].size_pt }); run_sizes_pt = $sizes; east_asia_fonts = $ea; latin_fonts = $lat
        bold_runs = @($runs | Where-Object { $_.SelectSingleNode('w:rPr/w:b[not(@w:val="0")]', $ns) }).Count; run_count = $runs.Count
        align = (A $p 'w:pPr/w:jc' 'val'); first_line_chars = (A $p 'w:pPr/w:ind' 'firstLineChars'); line = (A $p 'w:pPr/w:spacing' 'line'); numbered = [bool]$p.SelectSingleNode('w:pPr/w:numPr', $ns)
        ends_section = [bool]$p.SelectSingleNode('w:pPr/w:sectPr', $ns); has_drawing = [bool]$p.SelectSingleNode('.//w:drawing', $ns); has_math = [bool]$p.SelectSingleNode('.//m:oMath', $ns)
        text = $text }
}

$tables = @(); $t = 0
foreach ($tb in $doc.SelectNodes('/w:document/w:body/w:tbl', $ns)) {
    $t++
    $rows = @($tb.SelectNodes('w:tr', $ns))
    $tables += [ordered]@{ index = $t; locator = "/w:document[1]/w:body[1]/w:tbl[$t]"; rows = $rows.Count; columns = @($tb.SelectNodes('w:tblGrid/w:gridCol', $ns)).Count
        style_id = (A $tb 'w:tblPr/w:tblStyle' 'val'); border_top = (A $tb 'w:tblPr/w:tblBorders/w:top' 'val'); border_insideV = (A $tb 'w:tblPr/w:tblBorders/w:insideV' 'val')
        first_row_text = (($rows | Select-Object -First 1).SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join ' | ' }
}

$hf = @()
foreach ($name in ($parts.Keys | Where-Object { $_ -match '^word/(header|footer)\d+\.xml$' } | Sort-Object)) {
    $hd, $hns = Load $parts[$name]
    $hf += [ordered]@{ part = $name; text = (($hd.SelectNodes('//w:t', $hns) | ForEach-Object { $_.InnerText }) -join '')
        runs = @($hd.SelectNodes('//w:t', $hns) | ForEach-Object { $_.InnerText })
        fields = @($hd.SelectNodes('//w:instrText|//w:fldSimple', $hns) | ForEach-Object { if ($_.LocalName -eq 'fldSimple') { $_.GetAttribute('instr', $W) } else { $_.InnerText.Trim() } } | Where-Object { $_ }) }
}

# written formatting requirements inside the sample text (they beat the sample's own formatting)
$reqPattern = '(\d+(\.\d+)?\s*(pt|磅|点|号)|[一二三四五六七八小初]+号|黑体|宋体|仿宋|楷体|Times New Roman|Arial|bold|加粗|居中|centered|single[- ]spaced|行距|字号|font)'
$requirements = @($paras | Where-Object { $_.text -match $reqPattern } | ForEach-Object { [ordered]@{ locator = $_.locator; text = $_.text } })

$probe = [ordered]@{
    template = $TemplatePath; sha256 = (Get-FileHash $TemplatePath).Hash
    section_count = $sections.Count; sections = $sections
    paragraph_count = $paras.Count; paragraphs = $paras
    tables = $tables
    inline_drawings = @($doc.SelectNodes('//w:drawing/wp:inline', $ns)).Count; anchored_drawings = @($doc.SelectNodes('//w:drawing/wp:anchor', $ns)).Count
    omml_equations = @($doc.SelectNodes('//m:oMath', $ns)).Count; ole_objects = @($doc.SelectNodes('//w:object', $ns)).Count
    bookmarks = @($doc.SelectNodes('//w:bookmarkStart', $ns) | ForEach-Object { $_.GetAttribute('name', $W) })
    headers_footers = $hf
    written_requirements = $requirements
    styles_used = @($paras | Where-Object style_id | ForEach-Object { $styles[$_.style_id] } | Sort-Object { $_.id } -Unique)
}
$probe | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $OutDir 'template_probe.json') -Encoding utf8

$md = [Collections.Generic.List[string]]::new()
$md.Add("# Template probe: $(Split-Path $TemplatePath -Leaf)"); $md.Add(""); $md.Add("SHA-256 $($probe.sha256)")
$md.Add(""); $md.Add("## Sections ($($sections.Count))"); $md.Add("| # | locator | type | cols | page (twips) | margins t/r/b/l | titlePg |"); $md.Add("|---|---|---|---|---|---|---|")
foreach ($s in $sections) { $md.Add("| $($s.index) | ``$($s.locator)`` | $($s.type) | $($s.columns) | $($s.width_twips)x$($s.height_twips) | $($s.margins_twips.top)/$($s.margins_twips.right)/$($s.margins_twips.bottom)/$($s.margins_twips.left) | $($s.title_page) |") }
$md.Add(""); $md.Add("## Paragraphs ($($paras.Count))"); $md.Add("| p | style (size) | run sizes | fonts EA / Latin | bold | align | sect | text |"); $md.Add("|---|---|---|---|---|---|---|---|")
foreach ($p in $paras) {
    $txt = ($p.text -replace '\|', '/'); if ($txt.Length -gt 60) { $txt = $txt.Substring(0, 60) + '…' }
    $flags = @(); if ($p.ends_section) { $flags += 'END' }; if ($p.has_drawing) { $flags += 'IMG' }; if ($p.has_math) { $flags += 'MATH' }
    $md.Add("| $($p.index) | $($p.style_name) ($($p.style_size_pt)) | $($p.run_sizes_pt -join ',') | $($p.east_asia_fonts -join ',') / $($p.latin_fonts -join ',') | $($p.bold_runs)/$($p.run_count) | $($p.align) | $($flags -join ' ') | $txt |")
}
$md.Add(""); $md.Add("## Tables"); foreach ($tb in $tables) { $md.Add("- tbl[$($tb.index)] $($tb.rows)x$($tb.columns), top border $($tb.border_top), insideV $($tb.border_insideV): $($tb.first_row_text)") }
$md.Add(""); $md.Add("## Headers / footers"); foreach ($h in $hf) { $md.Add("- $($h.part): ``$($h.text)``  fields: $($h.fields -join '; ')  runs: $(($h.runs | ForEach-Object { "[$_]" }) -join '')") }
$md.Add(""); $md.Add("## Written requirements found in sample text"); foreach ($r in $requirements) { $md.Add("- $($r.locator): $($r.text)") }
$md.Add(""); $md.Add("Drawings: inline $($probe.inline_drawings), anchored $($probe.anchored_drawings); OMML $($probe.omml_equations); OLE $($probe.ole_objects); bookmarks $($probe.bookmarks.Count)")
$md | Set-Content (Join-Path $OutDir 'template_probe.md') -Encoding utf8
"probe written: $(Join-Path $OutDir 'template_probe.md') (+ .json)  sections=$($sections.Count) paragraphs=$($paras.Count) tables=$($tables.Count) written-requirements=$($requirements.Count)"
