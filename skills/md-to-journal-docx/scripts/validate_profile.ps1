<#
Check a journal profile (and optionally a paper's overrides) against its template, deterministically.
Every decision the agent wrote into format_mapping.json / profile.json is verified against the file.
Exit code 0 = no FAIL. WARN items need a look but do not block.

  pwsh -File validate_profile.ps1 -ProfileDir profiles/<name> [-PaperOverridesPath paper_overrides.json]
#>
param([Parameter(Mandatory)] [string]$ProfileDir, [string]$PaperOverridesPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem, System.Drawing
$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$ProfileDir = (Resolve-Path $ProfileDir).Path
$res = [Collections.Generic.List[object]]::new()
function Check($name, [bool]$ok, $detail, [switch]$Warn) { $res.Add([pscustomobject]@{ Status = $(if ($ok) { 'OK' } elseif ($Warn) { 'WARN' } else { 'FAIL' }); Check = $name; Detail = $detail }) }

$tpl = Join-Path $ProfileDir 'template_reference.docx'; $mapPath = Join-Path $ProfileDir 'format_mapping.json'; $profPath = Join-Path $ProfileDir 'profile.json'
foreach ($f in $tpl, $mapPath, $profPath) { Check "file $(Split-Path $f -Leaf)" (Test-Path $f) $f }
if ($res | Where-Object Status -eq 'FAIL') { $res | Format-Table -AutoSize -Wrap; exit 1 }
$map = Get-Content $mapPath -Raw -Encoding utf8 | ConvertFrom-Json -Depth 40
$prof = Get-Content $profPath -Raw -Encoding utf8 | ConvertFrom-Json -Depth 40

# merged rules, exactly as the builder merges them
function Merge($B, $O) { if ($null -eq $O) { return $B }; foreach ($k in @($O.Keys)) { if ($B.Contains($k) -and $B[$k] -is [Collections.IDictionary] -and $O[$k] -is [Collections.IDictionary]) { $B[$k] = Merge $B[$k] $O[$k] } else { $B[$k] = $O[$k] } }; return $B }
$rules = Get-Content (Join-Path $PSScriptRoot 'default_rules.json') -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
$rules = Merge $rules (Get-Content $profPath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable).rules
if ($PaperOverridesPath) { $rules = Merge $rules (Get-Content $PaperOverridesPath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable).rules }

# template parts
$zip = [IO.Compression.ZipFile]::OpenRead($tpl); $parts = @{}
try { foreach ($e in $zip.Entries) { if ($e.FullName -match '\.xml$') { $r = [IO.StreamReader]::new($e.Open()); try { $parts[$e.FullName] = $r.ReadToEnd() } finally { $r.Dispose() } } } } finally { $zip.Dispose() }
$doc = [Xml.XmlDocument]::new(); $doc.PreserveWhitespace = $true; $doc.LoadXml($parts['word/document.xml'])
$ns = [Xml.XmlNamespaceManager]::new($doc.NameTable); $ns.AddNamespace('w', $W); $ns.AddNamespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')

# 1. identity and approval
Check 'template hash = format_mapping.reference_sha256' ((Get-FileHash $tpl).Hash -eq $map.reference_sha256) "$((Get-FileHash $tpl).Hash) vs $($map.reference_sha256)"
Check 'authoring_gate = passed (template pages approved)' ($map.authoring_gate -eq 'passed') "authoring_gate=$($map.authoring_gate)"
Check 'template_visual_review.json present' (Test-Path (Join-Path $ProfileDir 'template_visual_review.json')) 'record of the page review' -Warn
$unres = @($map.unresolved | Where-Object { $_ }); Check 'format_mapping.unresolved is empty' ($unres.Count -eq 0) "$($unres.Count) open item(s): $(($unres | ForEach-Object { $_ | ConvertTo-Json -Compress -Depth 4 }) -join '; ')" -Warn

# 2. roles resolve to real paragraphs with the expected sample text
foreach ($prop in $map.role_mapping.PSObject.Properties) {
    $spec = $prop.Value; $node = $doc.SelectSingleNode([string]$spec.locator, $ns)
    $txt = if ($node) { ($node.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '' } else { '' }
    $ok = [bool]$node -and $node.LocalName -eq 'p'
    $detail = "$($spec.locator) -> '$(if ($txt.Length -gt 40) { $txt.Substring(0, 40) + '…' } else { $txt })'"
    if ($ok -and $spec.expected_text_pattern) { $ok = $txt -match [string]$spec.expected_text_pattern; $detail += " ~ /$($spec.expected_text_pattern)/" }
    Check "role $($prop.Name)" $ok $detail
    if ($node) { Check "role $($prop.Name) target size" ([double]$spec.target_size_pt -gt 0) "target_size_pt=$($spec.target_size_pt)" }
}
$roles = @($map.role_mapping.PSObject.Properties.Name)

# 3. every role the rules use exists in the mapping
$used = @($rules.front_matter.render_order) + @($rules.markdown.heading_roles.Values) + @($rules.body_roles.Values | Where-Object { $_ -notmatch '_secondary$' -or $roles -contains $_ })
foreach ($r in ($used | Sort-Object -Unique)) { Check "rules use role '$r'" ($roles -contains $r) 'must exist in format_mapping.role_mapping' }
foreach ($spec in @($rules.front_matter.roles)) {
    if ($spec.match) { $okRe = $true; try { [void][regex]::new([string]$spec.match) } catch { $okRe = $false }; Check "front-matter regex $($spec.role)" $okRe $spec.match }
    Check "front-matter role $($spec.role) is rendered" (@($rules.front_matter.render_order) -contains $spec.role) 'parsed roles not in render_order are silently dropped' -Warn
}
foreach ($k in 'body_start_pattern', 'references_heading_pattern', 'appendix_heading_pattern', 'table_note_pattern', 'citation_pattern') { $okRe = $true; try { [void][regex]::new([string]$rules.markdown[$k]) } catch { $okRe = $false }; Check "markdown.$k is a valid regex" $okRe $rules.markdown[$k] }

# 4. sections
$sect = @($doc.SelectNodes('//w:sectPr', $ns))
$sr = $rules.sections
Check 'rules.sections declared' ($null -ne $sr.front -and $null -ne $sr.body -and $null -ne $sr.tail) ($sr | ConvertTo-Json -Compress)
if ($null -ne $sr.body) {
    $all = @([int]$sr.front, [int]$sr.body) + @($sr.tail | ForEach-Object { [int]$_ })
    Check 'section indices exist in template' (@($all | Where-Object { $_ -ge $sect.Count }).Count -eq 0) "template has $($sect.Count) sections; declared $($all -join ',')"
    Check 'format_mapping.sections count = template sections' (@($map.sections).Count -eq $sect.Count) "$(@($map.sections).Count) vs $($sect.Count)"
    for ($i = 0; $i -lt [Math]::Min($sect.Count, @($map.sections).Count); $i++) {
        $cols = $sect[$i].SelectSingleNode('w:cols', $ns); $n = if ($cols -and $cols.GetAttribute('num', $W)) { [int]$cols.GetAttribute('num', $W) } else { 1 }
        Check "section $i columns" ([int]$map.sections[$i].columns -eq $n) "mapping $($map.sections[$i].columns) vs template $n"
    }
}

# 5. table / picture components
$tbl = $doc.SelectSingleNode([string]$map.tables.locator, $ns)
Check 'tables.locator is a table' ([bool]$tbl -and $tbl.LocalName -eq 'tbl') $map.tables.locator
Check 'template has an inline picture to clone' ([bool]$doc.SelectSingleNode('//w:drawing/wp:inline', $ns)) 'otherwise the builder uses a built-in minimal picture' -Warn
if ($prof.front_matter.classification) { Check 'classification locator resolves' ([bool]$doc.SelectSingleNode([string]$prof.front_matter.classification.locator, $ns)) $prof.front_matter.classification.locator }

# 6. header/footer fixes match exactly once
foreach ($fix in @($prof.header_fixes) + @($prof.footer_fixes)) {
    if (-not $fix) { continue }
    $xml = $parts[[string]$fix.part]
    if (-not $xml) { Check "fix part $($fix.part)" $false 'part not in template'; continue }
    $hd = [Xml.XmlDocument]::new(); $hd.LoadXml($xml); $hns = [Xml.XmlNamespaceManager]::new($hd.NameTable); $hns.AddNamespace('w', $W)
    $texts = @($hd.SelectNodes('//w:t', $hns) | ForEach-Object { $_.InnerText })
    if ($fix.replace_run_compact) { $n = @($texts | Where-Object { ($_ -replace '\s', '') -eq $fix.replace_run_compact }).Count; Check "header fix '$($fix.replace_run_compact)'" ($n -eq 1) "matches $n run(s) in $($fix.part)" }
    elseif ($fix.replace_span) { $n = [regex]::Matches(($texts -join ''), [regex]::Escape([string]$fix.replace_span)).Count; Check "header span '$($fix.replace_span)'" ($n -eq 1) "matches $n time(s) in $($fix.part)" }
    elseif ($fix.replace) {
        # replacements apply in order to each run (later pairs may match text produced by earlier ones)
        $work = @($texts)
        foreach ($pair in @($fix.replace)) {
            $n = @($work | Where-Object { $_.Contains([string]$pair[0]) }).Count
            Check "footer replace '$($pair[0])'" ($n -ge 1) "found in $n run(s) of $($fix.part) (after earlier replacements)" -Warn
            $work = @($work | ForEach-Object { $_.Replace([string]$pair[0], [string]$pair[1]) })
        }
    }
}

# 7. fonts
foreach ($f in @($prof.fonts_required)) { $ok = try { [void](New-Object System.Drawing.FontFamily $f); $true } catch { $false }; Check "font $f installed" $ok '' -Warn }

$res | Format-Table Status, Check, Detail -AutoSize -Wrap | Out-String -Width 220
$fail = @($res | Where-Object Status -eq 'FAIL').Count; $warn = @($res | Where-Object Status -eq 'WARN').Count
"checks: $($res.Count), FAIL: $fail, WARN: $warn"
if ($fail) { exit 1 }
