#requires -Version 7.0
<#
Markdown manuscript + journal profile -> editable DOCX with Word-native (OMML) equations.
No Word/WPS/COM is used here; native save, field update and page review are later steps.

  -ProfileDir      journal profile: template_reference.docx, format_mapping.json, profile.json
  -PaperConfigDir  per-paper inputs (all optional): figure_map.json, english_captions.json,
                   reference_translations.json, references_formatted.json, layout_config.json,
                   paper_overrides.json, source_numbering_correction.json
  -MathEngine      pandoc (default; texmath, broad LaTeX coverage) or cs (latex_to_omml.cs subset)
#>
param(
    [Parameter(Mandatory)] [string]$ManuscriptPath,
    [Parameter(Mandatory)] [string]$ProfileDir,
    [string]$PaperConfigDir = '',
    [string]$ReferencePath = '',
    [Parameter(Mandatory)] [string]$OutputPath,
    [string]$ReportDirectory = '',
    [string]$FigureMapPath = '',
    [string]$EnglishCaptionMapPath = '',
    [string]$ReferenceTranslationsPath = '',
    [string]$ReferenceFormattingPath = '',
    [string]$LayoutConfigPath = '',
    [string]$PaperOverridesPath = '',
    [string]$NumberingCorrectionPath = '',
    [ValidateSet('pandoc','cs')] [string]$MathEngine = 'pandoc',
    [string]$PythonPath = (Join-Path $env:LOCALAPPDATA 'md-to-journal-docx\venv\Scripts\python.exe'),
    [switch]$ValidateOnly,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing
if (-not ('ManuscriptTypesetting.Omml' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'latex_to_omml.cs') }
$ProfileDir = (Resolve-Path -LiteralPath $ProfileDir).Path
function Resolve-PaperInput([string]$Given,[string]$Name){
    if($Given){return $Given}
    if($PaperConfigDir){return (Join-Path $PaperConfigDir $Name)}
    return (Join-Path $env:TEMP ('md-to-journal-docx-absent-'+$Name))   # absent -> optional input skipped
}
if(-not $ReferencePath){$ReferencePath=Join-Path $ProfileDir 'template_reference.docx'}
$FigureMapPath=Resolve-PaperInput $FigureMapPath 'figure_map.json'
$EnglishCaptionMapPath=Resolve-PaperInput $EnglishCaptionMapPath 'english_captions.json'
$ReferenceTranslationsPath=Resolve-PaperInput $ReferenceTranslationsPath 'reference_translations.json'
$ReferenceFormattingPath=Resolve-PaperInput $ReferenceFormattingPath 'references_formatted.json'
$LayoutConfigPath=Resolve-PaperInput $LayoutConfigPath 'layout_config.json'
$PaperOverridesPath=Resolve-PaperInput $PaperOverridesPath 'paper_overrides.json'
$NumberingCorrectionPath=Resolve-PaperInput $NumberingCorrectionPath 'source_numbering_correction.json'
$JournalProfile=Get-Content -LiteralPath (Join-Path $ProfileDir 'profile.json') -Raw -Encoding utf8|ConvertFrom-Json -Depth 20
$Overrides=if(Test-Path -LiteralPath $PaperOverridesPath){Get-Content -LiteralPath $PaperOverridesPath -Raw -Encoding utf8|ConvertFrom-Json -Depth 20}else{[pscustomobject]@{}}
# Every convention lives in data, never in this script: default_rules.json <- profile.json "rules"
# <- paper_overrides.json "rules" (dictionaries merge per key; arrays/scalars replace). See references/rules.md.
function Merge-Rules($Base,$Over){
    if($null -eq $Over){return $Base}
    foreach($k in @($Over.Keys)){
        if($Base.Contains($k) -and $Base[$k] -is [Collections.IDictionary] -and $Over[$k] -is [Collections.IDictionary]){$Base[$k]=Merge-Rules $Base[$k] $Over[$k]}
        else{$Base[$k]=$Over[$k]}
    }
    return $Base
}
$Rules=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'default_rules.json') -Raw -Encoding utf8|ConvertFrom-Json -AsHashtable
$Rules=Merge-Rules $Rules (Get-Content -LiteralPath (Join-Path $ProfileDir 'profile.json') -Raw -Encoding utf8|ConvertFrom-Json -AsHashtable).rules
if(Test-Path -LiteralPath $PaperOverridesPath){$Rules=Merge-Rules $Rules (Get-Content -LiteralPath $PaperOverridesPath -Raw -Encoding utf8|ConvertFrom-Json -AsHashtable).rules}
$MdRules=$Rules.markdown;$FrontRules=$Rules.front_matter;$RoleNames=$Rules.body_roles
function Get-WordsPattern([object[]]$Words){return '(?:'+((@($Words)|ForEach-Object{[regex]::Escape([string]$_)}) -join '|')+')'}
$FigureWordsPattern=Get-WordsPattern $MdRules.figure_words
$TableWordsPattern=Get-WordsPattern $MdRules.table_words
$CaptionPattern='^\*\*((?<word>'+$FigureWordsPattern+'|'+$TableWordsPattern+')\s*(?<num>\d+)[^*]*)\*\*(?<note>.*)$'
$CaptionStartPattern='\*\*(?:'+$FigureWordsPattern+'|'+$TableWordsPattern+')\s*\d'
$FigureAltPattern='^'+$FigureWordsPattern+'\s*(\d+)'
$RefHang=[int]$Rules.references.hanging_indent_twips
$W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
$R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
$CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
$Utf8 = [Text.UTF8Encoding]::new($false)
$SourceFile = (Resolve-Path -LiteralPath $ManuscriptPath).Path
$ReferenceFile = (Resolve-Path -LiteralPath $ReferencePath).Path
$OutputFile = [IO.Path]::GetFullPath($OutputPath)
$OutputDirectory = Split-Path -LiteralPath $OutputFile
$ReportDirectoryPath = if($ReportDirectory){[IO.Path]::GetFullPath($ReportDirectory)}else{$OutputDirectory}
$SourceHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $SourceFile).Hash
$ReferenceHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $ReferenceFile).Hash
$Mapping = Get-Content -LiteralPath (Join-Path $ProfileDir 'format_mapping.json') -Raw | ConvertFrom-Json -Depth 30
if ($Mapping.reference_sha256 -ne $ReferenceHash) { throw 'Reference does not match the audited mapping.' }
if (-not $ValidateOnly -and $Mapping.authoring_gate -ne 'passed') { throw 'Template visual approval has not passed the content-fill gate.' }
if ($OutputFile.StartsWith($ProfileDir + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Output must not be written into the journal profile.' }
if ($OutputFile -eq $SourceFile -or $OutputFile -eq $ReferenceFile) { throw 'Output would overwrite an input.' }
if (-not $ValidateOnly -and (Test-Path -LiteralPath $OutputFile) -and -not $Force) { throw 'Output exists. Review it, then explicitly use -Force to replace it.' }
$Lines = [IO.File]::ReadAllLines($SourceFile, [Text.Encoding]::UTF8)
$Coverage = [Collections.Generic.HashSet[int]]::new()
$Blocks = [Collections.Generic.List[object]]::new()
$FormulaManifest = [Collections.Generic.List[object]]::new()
$ContentLedger = [Collections.Generic.List[object]]::new()
$ImageManifest = [Collections.Generic.List[object]]::new()
$ReferenceManifest = [Collections.Generic.List[object]]::new()
$TableLayoutManifest = [Collections.Generic.List[object]]::new()
$LayoutRelocationManifest = [Collections.Generic.List[object]]::new()
$AddedParts = @{}
$Replacements = @{}
$RemovedParts = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$RemovedRelationships = [Collections.Generic.List[object]]::new()
$Parts = @{}
$PartHashes = @{}
# Without layout_config.json every figure/table/equation stays in the column layout.
$FullWidthTables = @()
$FullWidthFigures = @()
$FullWidthEquations = @()
$CaptionLineSpacing = [double]$Rules.captions.line_spacing
$LayoutOptions = $null
if (Test-Path -LiteralPath $LayoutConfigPath) {
    $layout = Get-Content -LiteralPath $LayoutConfigPath -Raw | ConvertFrom-Json -Depth 15
    $LayoutOptions = $layout
    if ($null -ne $layout.full_width_tables) { $FullWidthTables = @($layout.full_width_tables) }
    if ($null -ne $layout.full_width_figures) { $FullWidthFigures = @($layout.full_width_figures) }
    if ($null -ne $layout.full_width_equations) { $FullWidthEquations = @($layout.full_width_equations) }
    if ($null -ne $layout.caption_line_spacing_multiplier) { $CaptionLineSpacing=[double]$layout.caption_line_spacing_multiplier }
    if($layout.revision_id -and $layout.outputs_must_stay_in_config_dir){
        $revisionRoot=[IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($LayoutConfigPath)).TrimEnd([IO.Path]::DirectorySeparatorChar)
        $revisionPrefix=$revisionRoot+[IO.Path]::DirectorySeparatorChar
        foreach($destination in @($OutputFile,$ReportDirectoryPath)){
            if($destination -ne $revisionRoot -and -not $destination.StartsWith($revisionPrefix,[StringComparison]::OrdinalIgnoreCase)){throw 'Revision outputs and reports must stay within the revision-config directory.'}
        }
    }
}
if($CaptionLineSpacing -le 0){throw 'Caption line spacing must be positive.'}
$EnglishCaptions=@{figures=@{};tables=@{}}
if(Test-Path -LiteralPath $EnglishCaptionMapPath){
    $captions=Get-Content -LiteralPath $EnglishCaptionMapPath -Raw|ConvertFrom-Json -Depth 10
    foreach($group in @('figures','tables')){
        foreach($item in $captions.$group){
            if(-not $item.number -or -not $item.title){throw 'English caption entries require number and title.'}
            $EnglishCaptions[$group][[int]$item.number]=[string]$item.title
        }
    }
}
$ReferenceTranslations=@{};$ReferenceFormatting=@{}
if(Test-Path -LiteralPath $ReferenceTranslationsPath){
    $translations=Get-Content -LiteralPath $ReferenceTranslationsPath -Raw|ConvertFrom-Json -Depth 15
    foreach($item in $translations.references){
        if(-not $item.number -or -not $item.original_chinese -or -not $item.english){throw 'Invalid bilingual reference entry.'}
        $ReferenceTranslations[[int]$item.number]=$item
    }
}
if(Test-Path -LiteralPath $ReferenceFormattingPath){
    $formatted=Get-Content -LiteralPath $ReferenceFormattingPath -Raw|ConvertFrom-Json -Depth 15
    foreach($item in $formatted.references){
        if(-not $item.number -or -not $item.source_exact -or -not $item.english){throw 'Formatted references require number, source_exact and english.'}
        if($ReferenceFormatting.ContainsKey([int]$item.number)){throw "Duplicate formatted reference $($item.number)"}
        $ReferenceFormatting[[int]$item.number]=$item
    }
}
$FigureMap = @{}
if (Test-Path -LiteralPath $FigureMapPath) {
    $figureConfig = Get-Content -LiteralPath $FigureMapPath -Raw | ConvertFrom-Json -Depth 20
    if ($null -eq $figureConfig.figures) { throw 'figure_map.json requires a figures array.' }
    foreach ($figure in $figureConfig.figures) {
        if (-not $figure.source -or -not $figure.path) { throw 'Every figure mapping requires source and path.' }
        if ($FigureMap.ContainsKey($figure.source)) { throw "Duplicate figure source: $($figure.source)" }
        $FigureMap[$figure.source] = $figure
    }
}

function Hash-Bytes([byte[]]$Bytes) {
    $hash = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($hash.ComputeHash($Bytes)).Replace('-','') } finally { $hash.Dispose() }
}
function Read-Xml([byte[]]$Bytes) {
    $xml = [Xml.XmlDocument]::new(); $xml.PreserveWhitespace = $true; $xml.XmlResolver = $null
    $xml.LoadXml([Text.Encoding]::UTF8.GetString($Bytes)); return ,$xml
}
function Xml-Bytes($Xml) {
    $stream = [IO.MemoryStream]::new(); $settings = [Xml.XmlWriterSettings]::new()
    $settings.Encoding = $Utf8; $settings.Indent = $false
    $writer = [Xml.XmlWriter]::Create($stream,$settings)
    try { $Xml.Save($writer); $writer.Flush(); return ,$stream.ToArray() } finally { $writer.Dispose(); $stream.Dispose() }
}
function New-W([string]$Name) { return ,$Doc.CreateElement('w',$Name,$W) }
function Set-W($Node,[string]$Name,$Value) { $null = $Node.SetAttribute($Name,$W,[string]$Value) }
function Ensure-W($Parent,[string]$Name) {
    $node = $Parent.SelectSingleNode("w:$Name",$Ns)
    if (-not $node) { $node = New-W $Name; $null = $Parent.AppendChild($node) }
    return ,$node
}
function Remove-Nodes($Node,[string]$XPath) { foreach ($child in @($Node.SelectNodes($XPath,$Ns))) { $null = $child.ParentNode.RemoveChild($child) } }
function Add-Block([string]$Kind,[string]$Text,[int]$Start,[int]$End,$Extra = $null) {
    for ($line=$Start; $line -le $End; $line++) { if ($Lines[$line].Trim()) { $null=$Coverage.Add($line) } }
    $block = [ordered]@{kind=$Kind;text=$Text;source_start_line=$Start+1;source_end_line=$End+1;source_block_id=('src_'+($Start+1)+'_'+($End+1)+'_'+$Kind);original_block_index=$Blocks.Count}
    if ($Extra) { foreach ($key in $Extra.Keys) { $block[$key]=$Extra[$key] } }
    $Blocks.Add($block)
}
function Parse-Markdown {
    $bodyStart = -1
    for ($i=0;$i -lt $Lines.Length;$i++) { if ($Lines[$i] -match $MdRules.body_start_pattern) { $bodyStart=$i;break } }
    if ($bodyStart -lt 0) { throw "No line matches rules.markdown.body_start_pattern ($($MdRules.body_start_pattern)): cannot tell front matter from body." }
    # Front matter: lines before the body start, matched against rules.front_matter.roles in list order;
    # roles with "unlabelled": n take the n-th line that matched no pattern.
    $front = [ordered]@{};$unlabelledSeen=0
    $patternRoles=@($FrontRules.roles|Where-Object{$_.match});$plainRoles=@($FrontRules.roles|Where-Object{$_.unlabelled}|Sort-Object{[int]$_.unlabelled})
    for ($i=0;$i -lt $bodyStart;$i++) {
        $line=$Lines[$i].Trim(); if (-not $line) { continue }
        $spec=$null;$content=$line
        foreach($candidate in $patternRoles){if($line -match $candidate.match){$spec=$candidate;$content=$Matches[1];break}}
        if(-not $spec){
            if($unlabelledSeen -ge $plainRoles.Count){throw "Front-matter line $($i+1) matches no rules.front_matter role: $line"}
            $spec=$plainRoles[$unlabelledSeen];$unlabelledSeen++
        }
        $role=[string]$spec.role
        if ($front.Contains($role)) { throw "Duplicate front-matter role $role (line $($i+1))" }
        $label=if($spec.label){[string]$spec.label}else{$null}
        $text=if($label){$label+[string]$FrontRules.label_separator+$content}else{$content}
        $front[$role]=[ordered]@{text=$text;label=$label;content=$content;source_line=$i+1};$null=$Coverage.Add($i)
    }
    foreach($spec in @($FrontRules.roles|Where-Object{$_.required})) {
        if(-not $front.Contains([string]$spec.role)){throw "Missing required front-matter role $($spec.role) (rules.front_matter.roles)"}
    }
    $script:Front=$front
    $i=$bodyStart
    while($i -lt $Lines.Length) {
        $line=$Lines[$i].Trim();if(-not $line){$i++;continue};$start=$i
        if($line -match '^(#{2,6})\s+(.+)$') {
            $level=$Matches[1].Length-1;$text=$Matches[2]
            $kind=if($text -match $MdRules.references_heading_pattern){'references_heading'}else{'heading'}
            Add-Block $kind $text $i $i @{level=$level};$i++;continue
        }
        if($line.StartsWith('$$')) {
            $formula=$line.Substring(2)
            while(-not $formula.TrimEnd().EndsWith('$$')) { $i++;if($i -ge $Lines.Length){throw 'Unclosed display formula'};$formula+=' '+$Lines[$i].Trim() }
            $formula=$formula.TrimEnd();$formula=$formula.Substring(0,$formula.Length-2).Trim()
            if(-not $formula){throw 'Empty display formula'}
            Add-Block 'equation' $formula $start $i;$i++;continue
        }
        if($line -match '^!\[([^\]]*)\]\(([^)]+)\)$') {
            Add-Block 'image' '' $i $i @{alt=$Matches[1];image_path=$Matches[2]};$i++;continue
        }
        if($line -match $CaptionPattern) {
            # copy the groups first: a second -match would overwrite $Matches
            $capText=$Matches[1];$capWord=$Matches['word'];$capNum=[int]$Matches['num'];$capNote=$Matches['note'].Trim()
            $kind=if($capWord -match ('^'+$FigureWordsPattern+'$')){'figure_caption'}else{'table_caption'}
            Add-Block $kind $capText $i $i @{number=$capNum;note=$capNote};$i++;continue
        }
        if($line.StartsWith('|')) {
            $rows=[Collections.Generic.List[object]]::new();$alignments=@();$tableLine=0
            while($i -lt $Lines.Length -and $Lines[$i].Trim().StartsWith('|')) {
                $cells=@($Lines[$i].Trim().Trim('|').Split('|')|ForEach-Object{$_.Trim()})
                if($tableLine -eq 1) {
                    foreach($cell in $cells){if($cell -notmatch '^:?-{3,}:?$'){throw "Invalid table separator on line $($i+1)"}}
                    $alignments=@($cells|ForEach-Object{if($_.EndsWith(':')){'right'}else{'left'}})
                }else{$rows.Add($cells)}
                $i++;$tableLine++
            }
            if($rows.Count -lt 2){throw "Table requires header and data on line $($start+1)"}
            $cols=$rows[0].Count;foreach($row in $rows){if($row.Count -ne $cols){throw "Nonrectangular table at line $($start+1)"}}
            Add-Block 'table' '' $start ($i-1) @{rows=$rows;alignments=$alignments};continue
        }
        if($line -match '^\[\d+\]\s+') { Add-Block 'reference' $line $i $i;$i++;continue }
        if($line -match '^\d+\.\s+' -or $line -match '^[-*]\s+') { Add-Block 'list_item' $line $i $i;$i++;continue }
        if($line.StartsWith('```')){throw "Code block has no approved template mapping on line $($i+1)"}
        $parts=[Collections.Generic.List[string]]::new();$parts.Add($line);$i++
        while($i -lt $Lines.Length) {
            $next=$Lines[$i].Trim()
            if(-not $next -or $next -match ('^(#{2,6}\s|!\[|\$\$|'+$CaptionStartPattern+'|\||\[\d+\]\s|\d+\.\s|[-*]\s)')){break}
            $parts.Add($next);$i++
        }
        Add-Block 'paragraph' ($parts -join ' ') $start ($i-1)
    }
    $missing=@(for($line=0;$line -lt $Lines.Length;$line++){if($Lines[$line].Trim() -and -not $Coverage.Contains($line)){$line+1}})
    if($missing.Count){throw "Unmapped source lines: $missing"}
    $tableNumber=0
    for($i=0;$i -lt $Blocks.Count;$i++){
        if($Blocks[$i].kind -eq 'table'){$tableNumber++;$Blocks[$i]['table_number']=$tableNumber}
        if($Blocks[$i].kind -eq 'paragraph' -and $i -gt 0 -and $Blocks[$i-1].kind -eq 'table' -and $Blocks[$i].text -match $MdRules.table_note_pattern){
            $Blocks[$i]['is_table_note']=$true;$Blocks[$i]['table_number']=$Blocks[$i-1].table_number
        }
    }
}

$zip=[IO.Compression.ZipFile]::OpenRead($ReferenceFile)
try {
    foreach($entry in $zip.Entries) {
        $stream=$entry.Open();$memory=[IO.MemoryStream]::new()
        try{$stream.CopyTo($memory);$bytes=$memory.ToArray()}finally{$memory.Dispose();$stream.Dispose()}
        $Parts[$entry.FullName]=$bytes;$PartHashes[$entry.FullName]=Hash-Bytes $bytes
    }
}finally{$zip.Dispose()}
$Doc=Read-Xml $Parts['word/document.xml']
$Ns=[Xml.XmlNamespaceManager]::new($Doc.NameTable)
foreach($pair in @(@('w',$W),@('m',$M),@('r',$R),@('wp','http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),@('a','http://schemas.openxmlformats.org/drawingml/2006/main'))){$Ns.AddNamespace($pair[0],$pair[1])}
$Body=$Doc.SelectSingleNode('/w:document/w:body',$Ns)
$SectionTemplates=@($Doc.SelectNodes('//w:sectPr',$Ns)|ForEach-Object{$_.CloneNode($true)})
# rules.sections declares which template section plays which part (front matter, body, closing sections).
$SecRules=$Rules.sections
if($null -eq $SecRules.front -or $null -eq $SecRules.body -or $null -eq $SecRules.tail){throw 'rules.sections (front/body/tail) is not declared in the profile; run validate_profile.ps1.'}
$SecFront=[int]$SecRules.front;$SecBody=[int]$SecRules.body;$SecTail=@($SecRules.tail|ForEach-Object{[int]$_})
$declared=@($SecFront,$SecBody)+$SecTail
if(@($declared|Where-Object{$_ -ge $SectionTemplates.Count -or $_ -lt 0}).Count -or $SecTail.Count -lt 1){throw "rules.sections refers to sections the template does not have (template has $($SectionTemplates.Count))."}
$ParagraphTemplates=@{}
foreach($property in $Mapping.role_mapping.PSObject.Properties){
    $source=$Doc.SelectSingleNode($property.Value.locator,$Ns)
    if(-not $source){throw "Missing source slot $($property.Name)"}
    $ParagraphTemplates[$property.Name]=$source.CloneNode($true)
}
$Classification=$JournalProfile.front_matter.classification   # optional generated placeholder line
if($Classification){
    $classificationSource=$Doc.SelectSingleNode([string]$Classification.locator,$Ns)
    if(-not $classificationSource){throw "Missing source classification-code paragraph $($Classification.locator)."}
    $ParagraphTemplates['classification_line']=$classificationSource.CloneNode($true)
}
$TableTemplate=$Doc.SelectSingleNode($Mapping.tables.locator,$Ns).CloneNode($true)
$templateDrawing=@($Doc.SelectNodes('//w:drawing[wp:inline]',$Ns))
if($templateDrawing.Count){$ImageTemplate=$templateDrawing[0].CloneNode($true)}else{
    # Template has no inline picture to clone: use a minimal standard inline picture.
    $fallback='<w:drawing xmlns:w="'+$W+'" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="'+$R+'"><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1" cy="1"/><wp:docPr id="1" name="figure"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="figure"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdPlaceholder"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
    $frag=[Xml.XmlDocument]::new();$frag.LoadXml($fallback);$ImageTemplate=$Doc.ImportNode($frag.DocumentElement,$true)
}
$DocumentRels=Read-Xml $Parts['word/_rels/document.xml.rels']
$ContentTypes=Read-Xml $Parts['[Content_Types].xml']
$BodySection=$Mapping.sections[$SecBody]
$BodyColumns=[Math]::Max(1,[int]$BodySection.columns)
$FullWidth=[int]$BodySection.width_twips-[int]$BodySection.margins_twips.left-[int]$BodySection.margins_twips.right
$ColumnWidth=[int][Math]::Floor(($FullWidth-[int]$BodySection.column_space_twips*($BodyColumns-1))/$BodyColumns)
$script:DrawingId=1000;$script:BookmarkId=1000

function New-Paragraph([string]$Role) {
    $p=$ParagraphTemplates[$Role].CloneNode($true)
    foreach($child in @($p.ChildNodes)){if($child.LocalName -ne 'pPr'){$null=$p.RemoveChild($child)}}
    $properties=Ensure-W $p 'pPr'
    Remove-Nodes $properties 'w:sectPr|w:rPr|w:contextualSpacing|w:pageBreakBefore'
    # Labels are supplied by Markdown; suppress inherited automatic numbering.
    $numbering=Ensure-W $properties 'numPr';$numId=Ensure-W $numbering 'numId';Set-W $numId 'val' '0'
    $widow=Ensure-W $properties 'widowControl';Set-W $widow 'val' '1'
    return ,$p
}
function Set-Paragraph($P,[string]$Align='both',[int]$FirstChars=0,[double]$Line=1.15,[bool]$KeepNext=$false,[bool]$PreserveSideIndents=$false) {
    $pr=Ensure-W $P 'pPr';$jc=Ensure-W $pr 'jc';Set-W $jc 'val' $Align
    $ind=Ensure-W $pr 'ind'
    foreach($attr in @('hanging','hangingChars')){ $ind.RemoveAttribute($attr,$W) }
    if(-not $PreserveSideIndents){foreach($attr in @('left','right','leftChars','rightChars')){$ind.RemoveAttribute($attr,$W)}}
    Set-W $ind 'firstLine' '0';Set-W $ind 'firstLineChars' $FirstChars
    $spacing=Ensure-W $pr 'spacing'
    if($Line -gt 0){Set-W $spacing 'line' ([int][Math]::Round($Line*240));Set-W $spacing 'lineRule' 'auto'}
    foreach($attr in @('beforeLines','afterLines','beforeAutospacing','afterAutospacing')){$spacing.RemoveAttribute($attr,$W)}
    $keep=Ensure-W $pr 'keepNext';Set-W $keep 'val' ([int]$KeepNext)
}
function Add-TextRun($P,[string]$Text,[string]$Role,[Nullable[bool]]$Bold=$null,[bool]$Italic=$false,[bool]$Superscript=$false,[double]$OverrideSize=0,[string]$EastAsiaFont='') {
    if($Text -eq ''){return}
    $spec=$Mapping.role_mapping.$Role
    $size=if($OverrideSize -gt 0){$OverrideSize}else{[double]$spec.target_size_pt}
    if($null -eq $Bold){$Bold=if($null -ne $spec.target_bold){[bool]$spec.target_bold}else{$false}}
    $run=New-W 'r';$rp=New-W 'rPr';$fonts=New-W 'rFonts'
    Set-W $fonts 'ascii' $spec.target_latin_font;Set-W $fonts 'hAnsi' $spec.target_latin_font;Set-W $fonts 'eastAsia' $(if($EastAsiaFont){$EastAsiaFont}else{$spec.target_east_asia_font});Set-W $fonts 'cs' $spec.target_latin_font
    $null=$rp.AppendChild($fonts)
    foreach($tag in @('b','bCs','i','iCs','color','sz','szCs','u')){
        $el=New-W $tag
        $val=switch($tag){'b'{[int][bool]$Bold};'bCs'{[int][bool]$Bold};'i'{[int]$Italic};'iCs'{[int]$Italic};'color'{'000000'};'u'{'none'};default{[int][Math]::Round($size*2)}}
        Set-W $el 'val' $val;$null=$rp.AppendChild($el)
    }
    if($Superscript){$v=New-W 'vertAlign';Set-W $v 'val' 'superscript';$null=$rp.AppendChild($v)}
    $null=$run.AppendChild($rp);$t=New-W 't';$t.InnerText=$Text;$null=$t.SetAttribute('space','http://www.w3.org/XML/1998/namespace','preserve');$null=$run.AppendChild($t);$null=$P.AppendChild($run)
}
$script:PandocCache=@{}
function Get-PandocOmml([object[]]$Requests){
    # One pandoc run for every not-yet-cached (tex, display) pair; results cached for Add-Formula.
    $todo=@($Requests|Where-Object{-not $script:PandocCache.ContainsKey($_.key)}|Sort-Object key -Unique)
    if(-not $todo.Count){return}
    if(-not(Test-Path -LiteralPath $PythonPath)){throw "MathEngine pandoc needs the Python env ($PythonPath); run scripts/check_runtime.ps1"}
    $in=Join-Path $env:TEMP ('tex2omml_in_'+[Guid]::NewGuid().ToString('N')+'.json');$out=$in.Replace('_in_','_out_')
    try{
        [IO.File]::WriteAllText($in,(ConvertTo-Json @($todo|ForEach-Object{[ordered]@{key=$_.key;tex=$_.tex;display=$_.display;size_pt=10.5}}) -Depth 4),$Utf8)
        & $PythonPath -I (Join-Path $PSScriptRoot 'tex2omml.py') --batch $in $out | Out-Null
        if(-not(Test-Path -LiteralPath $out)){throw 'tex2omml.py produced no output'}
        $res=Get-Content -LiteralPath $out -Raw -Encoding utf8|ConvertFrom-Json -AsHashtable
        foreach($k in $res.Keys){$script:PandocCache[$k]=$res[$k]}
    }finally{Remove-Item -LiteralPath $in,$out -ErrorAction SilentlyContinue}
}
function Convert-TexToOmml([string]$Tex,[bool]$Display,[int]$SourceLine,[double]$Size){
    if($MathEngine -eq 'cs'){return [ManuscriptTypesetting.Omml]::Convert($Doc,$Tex,$Size)}
    $key=$(if($Display){'D:'}else{'I:'})+$Tex
    Get-PandocOmml @([pscustomobject]@{key=$key;tex=$Tex;display=$Display})
    $hit=$script:PandocCache[$key]
    if($hit.error){throw $hit.error}
    # tex2omml.py writes 10.5 pt (w:sz 21); scale to the role's size
    $hp=[int][Math]::Round($Size*2)
    $xml=([string]$hit.omml).Replace('<w:sz w:val="21"/><w:szCs w:val="21"/>',"<w:sz w:val=`"$hp`"/><w:szCs w:val=`"$hp`"/>")
    $frag=[Xml.XmlDocument]::new();$frag.PreserveWhitespace=$true;$frag.LoadXml($xml)
    $node=$Doc.ImportNode($frag.DocumentElement,$true)
    # the document root already declares m:/w:; redundant xmlns on <m:oMath> breaks tools that match '<m:oMath>'
    foreach($a in @($node.Attributes|Where-Object{$_.Prefix -eq 'xmlns'})){$null=$node.Attributes.Remove($a)}
    return ,$node
}
function Get-ReadableTex([string]$Tex){
    try{return [ManuscriptTypesetting.Omml]::Readable($Tex)}catch{return $Tex}   # cs subset may not parse pandoc-only syntax
}
function Add-Formula($P,[string]$Tex,[bool]$Display,[int]$SourceLine,[double]$Size=10.5) {
    try{$math=Convert-TexToOmml $Tex $Display $SourceLine $Size}catch{throw "Formula conversion failed at source line ${SourceLine}: $($_.Exception.Message)"}
    $script:BookmarkId++;$name=if($Display){'display_equation_'+($FormulaManifest|Where-Object{$_.display}).Count.ToString('000')}else{'inline_math_'+$script:BookmarkId}
    $start=New-W 'bookmarkStart';Set-W $start 'id' $script:BookmarkId;Set-W $start 'name' $name
    $end=New-W 'bookmarkEnd';Set-W $end 'id' $script:BookmarkId
    $null=$P.AppendChild($start)
    if($Display){
        $mathParagraph=$Doc.CreateElement('m','oMathPara',$M);$mathPr=$Doc.CreateElement('m','oMathParaPr',$M);$mathJc=$Doc.CreateElement('m','jc',$M);$null=$mathJc.SetAttribute('val',$M,'center')
        $null=$mathPr.AppendChild($mathJc);$null=$mathParagraph.AppendChild($mathPr);$null=$mathParagraph.AppendChild($math);$null=$P.AppendChild($mathParagraph)
    }else{$null=$P.AppendChild($math)}
    $null=$P.AppendChild($end)
    $FormulaManifest.Add([ordered]@{source_line=$SourceLine;display=$Display;latex=$Tex;bookmark=$name;representation='editable_omml_intermediate';mathtype_verified=$false;readable_display=(Get-ReadableTex $Tex);omml_text=($math.SelectNodes('.//m:t',$Ns)|ForEach-Object{$_.InnerText}) -join '';omml_xml=$math.OuterXml})
}
function Add-Inline($P,[string]$Text,[string]$Role,[bool]$Citations=$true,[int]$SourceLine=0,[double]$OverrideSize=0) {
    if([regex]::Replace($Text,'\$[^$]+\$','') -match '(?<!\\)\$'){throw "Unmatched inline math delimiter on source line $SourceLine"}
    $pattern='\$[^$]+\$|\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`|'+$MdRules.citation_pattern
    if($MdRules.italic_token_pattern){$pattern+='|'+$MdRules.italic_token_pattern}
    $position=0
    foreach($match in [regex]::Matches($Text,$pattern)) {
        if($match.Index -gt $position){Add-TextRun $P $Text.Substring($position,$match.Index-$position) $Role -OverrideSize $OverrideSize}
        $token=$match.Value
        if($token.StartsWith('$')){Add-Formula $P $token.Substring(1,$token.Length-2) $false $SourceLine $(if($OverrideSize -gt 0){$OverrideSize}else{$Mapping.role_mapping.$Role.target_size_pt})}
        elseif($token.StartsWith('**')){Add-TextRun $P $token.Substring(2,$token.Length-4) $Role -Bold $true -OverrideSize $OverrideSize}
        elseif($token.StartsWith('*')){Add-TextRun $P $token.Substring(1,$token.Length-2) $Role -Italic $true -OverrideSize $OverrideSize}
        elseif($token.StartsWith('`')){Add-TextRun $P $token.Substring(1,$token.Length-2) $Role -OverrideSize $OverrideSize}
        elseif($token -match ('^'+$MdRules.citation_pattern+'$')){
            # rules.markdown.citation_superscript / citation_baseline_after_pattern (e.g. "see [3]" written in-line)
            $prefix=$Text.Substring(0,$match.Index)
            $sup=$Citations -and [bool]$MdRules.citation_superscript -and -not($MdRules.citation_baseline_after_pattern -and $prefix -match $MdRules.citation_baseline_after_pattern)
            Add-TextRun $P $token $Role -Superscript $sup -OverrideSize $OverrideSize
        }
        else{Add-TextRun $P $token $Role -Italic $true -OverrideSize $OverrideSize}   # rules.markdown.italic_token_pattern
        $position=$match.Index+$match.Length
    }
    if($position -lt $Text.Length){Add-TextRun $P $Text.Substring($position) $Role -OverrideSize $OverrideSize}
}
function Append-P($P){$null=$Body.AppendChild($P)}
function Add-Ledger($Block,$Node,[string]$Role) {
    $raw=($Lines[($Block.source_start_line-1)..($Block.source_end_line-1)] -join "`n")
    $ContentLedger.Add([ordered]@{kind=$Block.kind;role=$Role;source_block_id=$Block.source_block_id;source_start_line=$Block.source_start_line;source_end_line=$Block.source_end_line;source_raw=$raw;source_text=$Block.text;rendered_text=($Node.SelectNodes('.//w:t|.//m:t',$Ns)|ForEach-Object{$_.InnerText})-join '';layout_relocation=$Block.layout_relocation})
}
function Add-CaptionNote($Block,[string]$Text,[bool]$KeepNext=$false){
    if(-not $Text){return}
    $p=New-Paragraph $RoleNames.body;Set-Paragraph $p 'both' 0 $CaptionLineSpacing $KeepNext
    $keepLines=Ensure-W (Ensure-W $p 'pPr') 'keepLines';Set-W $keepLines 'val' 1
    Add-Inline $p $Text $RoleNames.body $true $Block.source_start_line ([double]$Rules.captions.note_size_pt);Append-P $p
    $ContentLedger.Add([ordered]@{kind='caption_note';role='caption_note';source_block_id=$Block.source_block_id;source_start_line=$Block.source_start_line;source_end_line=$Block.source_end_line;source_text=$Text;rendered_text=($p.SelectNodes('.//w:t|.//m:t',$Ns)|ForEach-Object{$_.InnerText})-join '';layout_relocation=$Block.layout_relocation})
}
function Apply-FigurePlacements {
    $beforeIds=@($Blocks|ForEach-Object{$_.source_block_id})
    $proseBefore=@($Blocks|Where-Object{$_.kind -notin @('image','figure_caption','table_caption','table') -and -not $_.is_table_note}|ForEach-Object{$_.source_block_id})
    $tablesBefore=@($Blocks|Where-Object{$_.kind -eq 'table'}|ForEach-Object{$_.table_number})
    foreach($placement in @($LayoutOptions.figure_placements)){
        if($null -eq $placement){continue}
        $number=[int]$placement.number;$indices=@(for($i=0;$i -lt $Blocks.Count;$i++){if($Blocks[$i].kind -eq 'image' -and $Blocks[$i].alt -match $FigureAltPattern -and [int]$Matches[1] -eq $number){$i}})
        if($indices.Count -ne 1){throw "Expected one image block for figure placement $number."}
        $sourceIndex=$indices[0]
        if($sourceIndex+1 -ge $Blocks.Count -or $Blocks[$sourceIndex+1].kind -ne 'figure_caption' -or $Blocks[$sourceIndex+1].number -ne $number){throw "Image/caption binding is not adjacent for figure $number."}
        $groupList=[Collections.Generic.List[object]]::new();$groupList.Add($Blocks[$sourceIndex]);$groupList.Add($Blocks[$sourceIndex+1])
        $removeIndices=[Collections.Generic.List[int]]::new();$removeIndices.Add($sourceIndex);$removeIndices.Add($sourceIndex+1)
        if($placement.include_table_number){
            $includedTable=[int]$placement.include_table_number
            $tableCaptions=@(for($j=0;$j -lt $Blocks.Count;$j++){if($Blocks[$j].kind -eq 'table_caption' -and $Blocks[$j].number -eq $includedTable){$j}})
            if($tableCaptions.Count -ne 1){throw "Expected one caption for included table $includedTable."}
            $tableIndex=$tableCaptions[0]
            if($tableIndex+1 -ge $Blocks.Count -or $Blocks[$tableIndex+1].kind -ne 'table' -or $Blocks[$tableIndex+1].table_number -ne $includedTable){throw 'Included table caption/body binding is not adjacent.'}
            foreach($j in @($tableIndex,($tableIndex+1))){$groupList.Add($Blocks[$j]);$removeIndices.Add($j)}
            if($tableIndex+2 -lt $Blocks.Count -and $Blocks[$tableIndex+2].is_table_note -and $Blocks[$tableIndex+2].table_number -eq $includedTable){$groupList.Add($Blocks[$tableIndex+2]);$removeIndices.Add($tableIndex+2)}
        }
        $group=[object[]]$groupList.ToArray()
        foreach($j in @($removeIndices|Sort-Object -Descending -Unique)){$Blocks.RemoveAt($j)}
        $anchors=@(for($i=0;$i -lt $Blocks.Count;$i++){if($Blocks[$i].kind -eq 'heading' -and $Blocks[$i].text -match $placement.after_heading_pattern){$i}})
        if($anchors.Count -ne 1){throw "Figure $number placement heading is missing or ambiguous."}
        $target=-1
        if($placement.after_paragraph_pattern){
            $paragraphAnchors=@(for($i=$anchors[0]+1;$i -lt $Blocks.Count;$i++){
                if($Blocks[$i].kind -in @('heading','references_heading')){break}
                if($Blocks[$i].kind -eq 'paragraph' -and -not $Blocks[$i].is_table_note -and $Blocks[$i].text -match $placement.after_paragraph_pattern){$i}
            })
            if($paragraphAnchors.Count -ne 1){throw "Figure $number paragraph pattern is missing or ambiguous within its heading."}
            $target=$paragraphAnchors[0]
        }else{
            $remaining=[int]$placement.after_paragraphs;if($remaining -lt 1){throw 'Figure placement after_paragraphs must be at least one.'}
            for($i=$anchors[0]+1;$i -lt $Blocks.Count;$i++){
                if($Blocks[$i].kind -in @('heading','references_heading')){break}
                if($Blocks[$i].kind -eq 'paragraph' -and -not $Blocks[$i].is_table_note){$remaining--;if($remaining -eq 0){$target=$i;break}}
            }
        }
        if($target -lt 0){throw "Figure $number placement has no matching paragraph anchor."}
        $relocationId='figure_'+$number+'_after_'+$Blocks[$target].source_block_id
        foreach($item in $group){$item['layout_relocation']=$relocationId}
        $LayoutRelocationManifest.Add([ordered]@{id=$relocationId;figure_number=$number;include_table_number=$placement.include_table_number;source_block_ids=@($group|ForEach-Object{$_.source_block_id});original_image_index=$group[0].original_block_index;after_heading=$Blocks[$anchors[0]].text;after_paragraphs=$placement.after_paragraphs;after_paragraph_pattern=$placement.after_paragraph_pattern;after_source_block_id=$Blocks[$target].source_block_id;after_source_line=$Blocks[$target].source_end_line;image_path=$group[0].image_path;caption_number=$group[1].number;reason='authorized layout-only placement of the complete bound visual group; table notes are excluded from scientific paragraph anchors'})
        $Blocks.InsertRange($target+1,$group)
    }
    $afterIds=@($Blocks|ForEach-Object{$_.source_block_id})
    if((($beforeIds|Sort-Object)-join '|') -cne (($afterIds|Sort-Object)-join '|')){throw 'Figure placement lost or duplicated a source block.'}
    $proseAfter=@($Blocks|Where-Object{$_.kind -notin @('image','figure_caption','table_caption','table') -and -not $_.is_table_note}|ForEach-Object{$_.source_block_id})
    if(($proseBefore-join '|') -cne ($proseAfter-join '|')){throw 'Visual placement changed scientific prose order.'}
    $tablesAfter=@($Blocks|Where-Object{$_.kind -eq 'table'}|ForEach-Object{$_.table_number})
    if(($tablesBefore-join ',') -cne ($tablesAfter-join ',')){throw 'Visual placement changed the order of tables.'}
}
function Insert-SoftBreakAfter($Paragraph,[string]$Marker){
    $textNodes=@($Paragraph.SelectNodes('w:r/w:t',$Ns));$plain=($textNodes|ForEach-Object{$_.InnerText})-join ''
    $first=$plain.IndexOf($Marker,[StringComparison]::Ordinal)
    if(-not $Marker -or $first -lt 0 -or $plain.IndexOf($Marker,$first+$Marker.Length,[StringComparison]::Ordinal) -ge 0){throw "Table soft-break marker is missing or ambiguous: $Marker"}
    $position=$first+$Marker.Length;$offset=0
    foreach($textNode in $textNodes){
        $value=$textNode.InnerText;$end=$offset+$value.Length
        if($position -gt $offset -and $position -le $end){
            $local=$position-$offset;$textNode.InnerText=$value.Substring(0,$local)
            $br=New-W 'br';$null=$textNode.ParentNode.InsertAfter($br,$textNode)
            if($local -lt $value.Length){$suffix=New-W 't';$suffix.InnerText=$value.Substring($local);$null=$suffix.SetAttribute('space','http://www.w3.org/XML/1998/namespace','preserve');$null=$textNode.ParentNode.InsertAfter($suffix,$br)}
            return
        }
        $offset=$end
    }
    throw 'Table soft-break position could not be located.'
}
function Add-Reference($Block){
    if($Block.text -notmatch '^\[(\d+)\]\s+(.+)$'){throw 'Invalid reference label.'}
    $number=[int]$Matches[1];$source=$Matches[2];$english=$source;$chinese=$null;$origin='source Markdown';$notes=@()
    if($ReferenceFormatting.ContainsKey($number)){
        $item=$ReferenceFormatting[$number]
        if($item.source_exact -cne $source){throw "Formatted reference source no longer matches [$number]."}
        $english=$item.english;$chinese=$item.chinese;$origin='references_formatted.json';$notes=$item.notes
    }elseif($ReferenceTranslations.ContainsKey($number)){
        $item=$ReferenceTranslations[$number]
        if($item.original_chinese -cne $source){throw "Bilingual reference source no longer matches [$number]."}
        $english=$item.english;$chinese=$source;$origin='reference_translations.json';$notes=$item.notes
    }
    $p=New-Paragraph 'reference_entry';Set-Paragraph $p 'both' 0 1 ([bool]$chinese)
    $ind=Ensure-W (Ensure-W $p 'pPr') 'ind';$ind.RemoveAttribute('firstLineChars',$W);$ind.RemoveAttribute('firstLine',$W);Set-W $ind 'left' $RefHang;Set-W $ind 'hanging' $RefHang
    Add-Inline $p ("[$number] "+$english) 'reference_entry' $false $Block.source_start_line
    Append-P $p;Add-Ledger $Block $p 'reference_entry'
    if($chinese){
        $cp=New-Paragraph 'reference_entry';Set-Paragraph $cp 'both' 0 1 $false
        $ind=Ensure-W (Ensure-W $cp 'pPr') 'ind';$ind.RemoveAttribute('firstLineChars',$W);Set-W $ind 'firstLine' 0;Set-W $ind 'left' $RefHang
        Add-Inline $cp $chinese 'reference_entry' $false $Block.source_start_line;Append-P $cp
        $ContentLedger.Add([ordered]@{kind='reference_chinese_counterpart';role='reference_entry';reference_number=$number;source_start_line=$Block.source_start_line;source_end_line=$Block.source_end_line;source_text=$source;rendered_text=($cp.SelectNodes('.//w:t|.//m:t',$Ns)|ForEach-Object{$_.InnerText})-join '';source=$origin})
    }
    $ReferenceManifest.Add([ordered]@{number=$number;source_exact=$source;english=$english;chinese=$chinese;source=$origin;notes=$notes;label_count=1})
}
function Add-SectionEnd($Template,[int]$Columns,[bool]$Inserted=$false) {
    $p=New-Paragraph $RoleNames.body;Set-Paragraph $p 'left' 0 1 $false
    $pr=Ensure-W $p 'pPr';$spacing=Ensure-W $pr 'spacing';Set-W $spacing 'before' 0;Set-W $spacing 'after' 0;Set-W $spacing 'line' 20;Set-W $spacing 'lineRule' 'exact'
    $section=$Template.CloneNode($true);$cols=$section.SelectSingleNode('w:cols',$Ns);Set-W $cols 'num' $Columns
    if($Inserted){
        # Added continuous layout sections must not restart first-page metadata/footer behavior.
        Remove-Nodes $section 'w:titlePg|w:pgNumType'
    }
    $null=$pr.AppendChild($section);Append-P $p
}
function Set-WidthMode([string]$Mode) {
    $new=if($Mode -eq 'full'){1}else{$BodyColumns}
    if($script:CurrentColumns -ne $new){Add-SectionEnd $SectionTemplates[$SecBody] $script:CurrentColumns $true;$script:CurrentColumns=$new}
}
function New-Table($Block,[int]$Number,[string]$Mode) {
    $table=$TableTemplate.CloneNode($true);Remove-Nodes $table 'w:tr'
    $pr=$table.SelectSingleNode('w:tblPr',$Ns);$layout=Ensure-W $pr 'tblLayout';Set-W $layout 'type' 'fixed'
    $total=if($Mode -eq 'full'){$FullWidth}else{$ColumnWidth}
    $rows=$Block.rows;$n=$rows[0].Count
    # Equal columns unless layout_config.json table_layouts gives column_weights for this table.
    $weights=@(1..$n|ForEach-Object{1})
    $tableOptions=$null
    if($LayoutOptions.table_layouts){$property=$LayoutOptions.table_layouts.PSObject.Properties[[string]$Number];if($property){$tableOptions=$property.Value}}
    if($tableOptions.column_weights){$weights=@($tableOptions.column_weights)}
    if($weights.Count -ne $n){$weights=@(1..$n|ForEach-Object{1})}
    if(@($weights|Where-Object{[double]$_ -le 0}).Count){throw "Table $Number has an invalid column weight."}
    $sideMargin=if($Mode -eq 'column'){[int]$Rules.tables.cell_side_margin_twips.column}else{[int]$Rules.tables.cell_side_margin_twips.full}
    $padTB=[int]$Rules.tables.cell_padding_top_bottom_twips;$ruleOuter=[int]$Rules.tables.rule_sizes_eighth_pt.outer;$ruleHeader=[int]$Rules.tables.rule_sizes_eighth_pt.header
    $tableFont=if($null -ne $Rules.tables.font_size_pt){[double]$Rules.tables.font_size_pt}else{[double]$Mapping.tables.body_size_pt}
    if($null -ne $tableOptions.cell_side_margin_twips){$sideMargin=[int]$tableOptions.cell_side_margin_twips}
    if($sideMargin -lt 0){throw 'Cell side margin cannot be negative.'}
    $weightSum=($weights|Measure-Object -Sum).Sum;$widths=@($weights|ForEach-Object{[int][Math]::Floor($total*$_/$weightSum)});$widths[0]+=$total-($widths|Measure-Object -Sum).Sum
    $tblW=Ensure-W $pr 'tblW';Set-W $tblW 'type' 'dxa';Set-W $tblW 'w' $total
    $borders=Ensure-W $pr 'tblBorders'
    foreach($side in @('top','left','bottom','right','insideH','insideV')){$b=Ensure-W $borders $side;Set-W $b 'val' $(if($side -in @('top','bottom')){'single'}else{'nil'});Set-W $b 'sz' $(if($side -in @('top','bottom')){$ruleOuter}else{0});Set-W $b 'color' '000000';Set-W $b 'space' 0}
    $grid=$table.SelectSingleNode('w:tblGrid',$Ns);$grid.RemoveAll()
    foreach($width in $widths){$col=New-W 'gridCol';Set-W $col 'w' $width;$null=$grid.AppendChild($col)}
    $sourceRow=$TableTemplate.SelectSingleNode('w:tr[1]',$Ns);$sourceCell=$sourceRow.SelectSingleNode('w:tc[1]',$Ns)
    for($rowIndex=0;$rowIndex -lt $rows.Count;$rowIndex++){
        $row=$sourceRow.CloneNode($true);Remove-Nodes $row 'w:tc'
        $trPr=Ensure-W $row 'trPr';Remove-Nodes $trPr 'w:trHeight|w:tblHeader';$cant=Ensure-W $trPr 'cantSplit';Set-W $cant 'val' 1
        if($rowIndex -eq 0){$repeat=Ensure-W $trPr 'tblHeader';Set-W $repeat 'val' 1}
        for($colIndex=0;$colIndex -lt $n;$colIndex++){
            $cell=$sourceCell.CloneNode($true);foreach($child in @($cell.ChildNodes)){if($child.LocalName -ne 'tcPr'){$null=$cell.RemoveChild($child)}}
            $tcPr=Ensure-W $cell 'tcPr';Remove-Nodes $tcPr 'w:gridSpan|w:vMerge|w:shd|w:noWrap|w:tcFitText'
            $tcW=Ensure-W $tcPr 'tcW';Set-W $tcW 'type' 'dxa';Set-W $tcW 'w' $widths[$colIndex]
            $vertical=Ensure-W $tcPr 'vAlign';Set-W $vertical 'val' 'center'
            $margin=Ensure-W $tcPr 'tcMar';foreach($side in @('top','left','bottom','right')){$edge=Ensure-W $margin $side;Set-W $edge 'type' 'dxa';Set-W $edge 'w' $(if($side -in @('top','bottom')){$padTB}else{$sideMargin})}
            $cb=Ensure-W $tcPr 'tcBorders'
            foreach($side in @('top','left','bottom','right')){$edge=Ensure-W $cb $side;$visible=($rowIndex -eq 0 -and $side -in @('top','bottom'))-or($rowIndex -eq $rows.Count-1 -and $side -eq 'bottom');Set-W $edge 'val' $(if($visible){'single'}else{'nil'});Set-W $edge 'sz' $(if($rowIndex -eq 0 -and $side -eq 'bottom'){$ruleHeader}elseif($visible){$ruleOuter}else{0});Set-W $edge 'color' '000000'}
            $p=New-Paragraph $RoleNames.body;$align=if($rowIndex -eq 0){'center'}elseif($Block.alignments[$colIndex] -eq 'right'){'right'}else{'left'}
            Set-Paragraph $p $align 0 1 ($rowIndex -lt $rows.Count-1);$spacing=Ensure-W (Ensure-W $p 'pPr') 'spacing';Set-W $spacing 'before' 0;Set-W $spacing 'after' 0
            Add-Inline $p ([string]$rows[$rowIndex][$colIndex]) $RoleNames.body $true $Block.source_start_line $tableFont
            foreach($softBreak in @($tableOptions.soft_breaks)){
                if($null -ne $softBreak -and [int]$softBreak.row -eq $rowIndex+1 -and [int]$softBreak.column -eq $colIndex+1){Insert-SoftBreakAfter $p ([string]$softBreak.after)}
            }
            if($rowIndex -eq 0){foreach($b in $p.SelectNodes('w:r/w:rPr/w:b|w:r/w:rPr/w:bCs',$Ns)){Set-W $b 'val' 1}}
            $null=$cell.AppendChild($p);$null=$row.AppendChild($cell)
        }
        $null=$table.AppendChild($row)
    }
    # Verify soft breaks and width changes did not alter any cell's text or numbers.
    $generatedRows=$table.SelectNodes('w:tr',$Ns)
    for($rowIndex=0;$rowIndex -lt $rows.Count;$rowIndex++){
        $cells=$generatedRows[$rowIndex].SelectNodes('w:tc',$Ns)
        for($colIndex=0;$colIndex -lt $n;$colIndex++){
            $expected=([string]$rows[$rowIndex][$colIndex]) -replace '\*\*([^*]+)\*\*','$1' -replace '\*([^*]+)\*','$1' -replace '`([^`]+)`','$1'
            if($expected.Contains('$')){continue} # Math retention is checked through its OMML/source manifest.
            $actual=($cells[$colIndex].SelectNodes('.//w:t',$Ns)|ForEach-Object{$_.InnerText})-join ''
            if($actual -cne $expected){throw "Table $Number cell text changed at row $($rowIndex+1), column $($colIndex+1)."}
        }
    }
    $TableLayoutManifest.Add([ordered]@{number=$Number;source_block_id=$Block.source_block_id;width_mode=$Mode;width_twips=$total;width_pt=$total/20;column_widths_twips=$widths;column_weights=$weights;side_margin_twips=$sideMargin;font_size_pt=$tableFont;line_spacing_twips=240;soft_breaks=@($tableOptions.soft_breaks|Where-Object{$null -ne $_});cell_text_preserved=$true;keep_table_together=$true})
    return ,$table
}
function Add-Image($Block,[string]$Mode) {
    $path=$Block.image_path;$effective=$null
    if($FigureMap.ContainsKey($path)){$effective=$FigureMap[$path];$path=$effective.path;$base=Split-Path -LiteralPath ([IO.Path]::GetFullPath($FigureMapPath))}else{$base=Split-Path -LiteralPath $SourceFile}
    $file=if([IO.Path]::IsPathRooted($path)){[IO.Path]::GetFullPath($path)}else{[IO.Path]::GetFullPath((Join-Path $base $path))}
    if(-not(Test-Path -LiteralPath $file -PathType Leaf)){throw "Missing image: $file"}
    $bytes=[IO.File]::ReadAllBytes($file);$memory=[IO.MemoryStream]::new($bytes,$false)
    $image=[Drawing.Image]::FromStream($memory)
    try{$pixelWidth=$image.Width;$pixelHeight=$image.Height}finally{$image.Dispose();$memory.Dispose()}
    $extension=[IO.Path]::GetExtension($file).TrimStart('.').ToLowerInvariant();if($extension -notin @('png','jpg','jpeg')){throw "Unsupported image extension $extension"}
    $number=$ImageManifest.Count+1;$part="word/media/manuscript_figure_$number.$extension";$AddedParts[$part]=$bytes
    $rid="rIdManuscriptFigure$number"
    if($DocumentRels.DocumentElement.ChildNodes|Where-Object{$_.GetAttribute('Id') -eq $rid}){throw "Existing image relationship $rid"}
    $rel=$DocumentRels.CreateElement('Relationship',$PR);$rel.SetAttribute('Id',$rid);$rel.SetAttribute('Type',$R+'/image');$rel.SetAttribute('Target',$part.Substring(5));$null=$DocumentRels.DocumentElement.AppendChild($rel)
    if(-not($ContentTypes.DocumentElement.ChildNodes|Where-Object{$_.LocalName -eq 'Default' -and $_.GetAttribute('Extension') -eq $extension})){$default=$ContentTypes.CreateElement('Default',$CT);$default.SetAttribute('Extension',$extension);$default.SetAttribute('ContentType',$(if($extension -eq 'png'){'image/png'}else{'image/jpeg'}));$null=$ContentTypes.DocumentElement.AppendChild($default)}
    $widthTwips=if($Mode -eq 'full'){$FullWidth}else{$ColumnWidth};$cx=[long]$widthTwips*635;$cy=[long][Math]::Round($cx*$pixelHeight/$pixelWidth)
    # Limit a figure to the usable page height, preserving its aspect ratio.
    $maxCy=[long]([int]$BodySection.height_twips-[int]$BodySection.margins_twips.top-[int]$BodySection.margins_twips.bottom-1200)*635
    if($cy -gt $maxCy){$cx=[long][Math]::Round($cx*$maxCy/$cy);$cy=$maxCy}
    $p=New-Paragraph $RoleNames.body;Set-Paragraph $p 'center' 0 1 $true
    $run=New-W 'r';$drawing=$ImageTemplate.CloneNode($true)
    $extent=$drawing.SelectSingleNode('wp:inline/wp:extent',$Ns);$extent.SetAttribute('cx',[string]$cx);$extent.SetAttribute('cy',[string]$cy)
    foreach($ext in $drawing.SelectNodes('.//a:xfrm/a:ext',$Ns)){$ext.SetAttribute('cx',[string]$cx);$ext.SetAttribute('cy',[string]$cy)}
    $script:DrawingId++;$dp=$drawing.SelectSingleNode('wp:inline/wp:docPr',$Ns);$dp.SetAttribute('id',[string]$script:DrawingId);$dp.SetAttribute('name',$Block.alt);$dp.SetAttribute('descr',$Block.alt)
    $blip=$drawing.SelectSingleNode('.//a:blip',$Ns);$null=$blip.SetAttribute('embed',$R,$rid)
    foreach($sourceRect in @($drawing.SelectNodes('.//a:srcRect',$Ns))){$null=$sourceRect.ParentNode.RemoveChild($sourceRect)}
    $null=$run.AppendChild($drawing);$null=$p.AppendChild($run);Append-P $p
    $ImageManifest.Add([ordered]@{source=$Block.image_path;path=$file;sha256=Hash-Bytes $bytes;alt=$Block.alt;width_mode=$Mode;pixel_width=$pixelWidth;pixel_height=$pixelHeight;display_width_mm=[Math]::Round($cx/36000,3);display_height_mm=[Math]::Round($cy/36000,3);part=$part;relationship_id=$rid})
    Add-Ledger $Block $p 'image'
}

Parse-Markdown
if($MathEngine -eq 'pandoc'){
    # prefetch every formula in one pandoc run (Add-Formula falls back to single calls for misses)
    # NB: PowerShell names are case-insensitive - never reuse $W/$M/$R/$PR/$CT (namespace URIs) as loop variables
    $prefetch=[Collections.Generic.List[object]]::new()
    $prefetchTexts=@($Front.Values|ForEach-Object{$_.text})+@($Blocks|Where-Object{$_.kind -ne 'equation'}|ForEach-Object{$_.text})+@($Blocks|Where-Object{$_.rows}|ForEach-Object{$_.rows|ForEach-Object{$_}})
    foreach($prefetchText in $prefetchTexts){foreach($prefetchMatch in [regex]::Matches([string]$prefetchText,'\$[^$]+\$')){$prefetchTex=$prefetchMatch.Value.Substring(1,$prefetchMatch.Value.Length-2);$prefetch.Add([pscustomobject]@{key='I:'+$prefetchTex;tex=$prefetchTex;display=$false})}}
    foreach($prefetchBlock in @($Blocks|Where-Object{$_.kind -eq 'equation'})){$prefetch.Add([pscustomobject]@{key='D:'+$prefetchBlock.text;tex=$prefetchBlock.text;display=$true})}
    Get-PandocOmml $prefetch.ToArray()
    # report every unconvertible formula at once, with its source line(s)
    $badFormulas=@($prefetch|Sort-Object key -Unique|Where-Object{$script:PandocCache[$_.key].error})
    if($badFormulas.Count){
        $msg=@($badFormulas|ForEach-Object{$tex=$_.tex;$lineNos=@(for($n=0;$n -lt $Lines.Length;$n++){if($Lines[$n].Contains($tex)){$n+1}});"  line $($lineNos -join ','): $tex"})
        throw ("$($badFormulas.Count) formula(s) pandoc cannot convert (fix the LaTeX, or try -MathEngine cs):`n"+($msg -join "`n"))
    }
}
Apply-FigurePlacements
foreach($child in @($Body.ChildNodes)){$null=$Body.RemoveChild($child)}
# Front matter in rules.front_matter.render_order; roles absent from the Markdown get missing_text.
$RoleSpecs=@{};foreach($spec in @($FrontRules.roles)){$RoleSpecs[[string]$spec.role]=$spec}
foreach($role in @($FrontRules.render_order)){
    $role=[string]$role
    if(-not $ParagraphTemplates.ContainsKey($role)){throw "Front-matter role $role has no format_mapping.json role_mapping entry."}
    $render=if($FrontRules.render -and $FrontRules.render.Contains($role)){$FrontRules.render[$role]}else{@{}}
    $p=New-Paragraph $role
    Set-Paragraph $p $(if($render.align){[string]$render.align}else{'both'}) 0 $(if($null -ne $render.line_spacing){[double]$render.line_spacing}else{0}) ([bool]$render.keep_next) $true
    if($render.suppress_hyphenation){$hyphen=Ensure-W (Ensure-W $p 'pPr') 'suppressAutoHyphens';Set-W $hyphen 'val' 1}
    $entry=if($Front.Contains($role)){$Front[$role]}else{$null}
    $text=if($entry){$entry.text}else{[string]$FrontRules.missing_text}
    $line=if($entry){$entry.source_line}else{0}
    if($entry -and $entry.label){
        $spec=$RoleSpecs[$role]
        Add-TextRun $p ($entry.label+[string]$FrontRules.label_separator) $role -Bold $true -EastAsiaFont $(if($spec.label_east_asia_font){[string]$spec.label_east_asia_font}else{''});Add-Inline $p $entry.content $role $true $line
    }elseif($role -eq [string]$FrontRules.document_title_role -and $Overrides.title_break_before -and $text.Contains([string]$Overrides.title_break_before)){
        # paper_overrides.json title_break_before: layout-only line break in the document title
        $breakAt=$text.IndexOf([string]$Overrides.title_break_before);Add-Inline $p $text.Substring(0,$breakAt) $role $false $line
        $breakRun=New-W 'r';$breakNode=New-W 'br';$null=$breakRun.AppendChild($breakNode);$null=$p.AppendChild($breakRun)
        Add-Inline $p $text.Substring($breakAt) $role $false $line
    }else{Add-Inline $p $text $role $false $line}
    Append-P $p;$ContentLedger.Add([ordered]@{kind='front_matter';role=$role;source_start_line=$line;source_end_line=$line;source_text=$text;rendered_text=($p.SelectNodes('.//w:t|.//m:t',$Ns)|ForEach-Object{$_.InnerText})-join ''})
    if($Classification -and $role -eq [string]$Classification.after_role){
        $classText=[string]$Classification.text
        # NB: not $classification - PowerShell names are case-insensitive and $Classification is the rule object
        $classPara=New-Paragraph 'classification_line';Set-Paragraph $classPara 'both' 0 1 $false $true
        Add-TextRun $classPara $classText ([string]$Classification.after_role) -Bold $false;Append-P $classPara
        $ContentLedger.Add([ordered]@{kind='generated_placeholder';role='classification_line';source_start_line=0;source_end_line=0;source_text=$classText;rendered_text=$classText;template_locator=[string]$Classification.locator;reason='classification code not supplied; do not infer one'})
    }
}
Add-SectionEnd $SectionTemplates[$SecFront] ([Math]::Max(1,[int]$Mapping.sections[$SecFront].columns))
$script:CurrentColumns=$BodyColumns;$tableNumber=0;$equationNumber=0;$referencesStarted=$false;$CaptionModes=@{};$PendingTableNotes=@{}
$ParaRules=$Rules.paragraphs
for($index=0;$index -lt $Blocks.Count;$index++){
    $block=$Blocks[$index]
    $wide=$false;$mode='column'
    if($block.kind -eq 'image'){
        if($block.alt -match $FigureAltPattern){$wide=[int]$Matches[1] -in $FullWidthFigures}
        if($FigureMap.ContainsKey($block.image_path) -and $FigureMap[$block.image_path].width_mode){$wide=$FigureMap[$block.image_path].width_mode -eq 'full'}
        if($block.alt -match $FigureAltPattern){$CaptionModes[[int]$Matches[1]]=$(if($wide){'full'}else{'column'})}
    }elseif($block.kind -in @('figure_caption','table_caption')){
        $wide=if($block.kind -eq 'figure_caption'){$block.number -in $FullWidthFigures}else{$block.number -in $FullWidthTables}
        if($block.kind -eq 'figure_caption' -and $CaptionModes.ContainsKey($block.number)){$wide=$CaptionModes[$block.number] -eq 'full'}
    }elseif($block.kind -eq 'table'){$wide=($tableNumber+1) -in $FullWidthTables}
    elseif($block.kind -eq 'equation'){$equationNumber++;$wide=$equationNumber -in $FullWidthEquations}
    elseif($block.kind -eq 'paragraph' -and $block.text -match $MdRules.table_note_pattern -and $index -gt 0 -and $Blocks[$index-1].kind -eq 'table'){$wide=$tableNumber -in $FullWidthTables}
    if($wide){$mode='full'}
    Set-WidthMode $mode
    if($block.kind -eq 'image'){Add-Image $block $mode;continue}
    if($block.kind -eq 'table'){
        $tableNumber++;$table=New-Table $block $tableNumber $mode;$null=$Body.AppendChild($table);Add-Ledger $block $table 'table'
        if($PendingTableNotes.ContainsKey($tableNumber)){Add-CaptionNote $PendingTableNotes[$tableNumber].block $PendingTableNotes[$tableNumber].text}
        continue
    }
    if($block.kind -eq 'reference'){Add-Reference $block;continue}
    $role=switch($block.kind){
        'heading'{$levels=$MdRules.heading_roles;$key=[string]$block.level;if($levels.Contains($key)){[string]$levels[$key]}else{[string]$levels[[string](@($levels.Keys|ForEach-Object{[int]$_}|Measure-Object -Maximum).Maximum)]}}
        'references_heading'{$RoleNames.references_heading};'reference'{$RoleNames.reference_entry}
        'figure_caption'{$RoleNames.figure_caption};'table_caption'{$RoleNames.table_caption};default{$RoleNames.body}}
    if(-not $ParagraphTemplates.ContainsKey($role)){throw "Role '$role' (block $($block.source_block_id)) has no format_mapping.json role_mapping entry."}
    $p=New-Paragraph $role
    $heading=$block.kind -in @('heading','references_heading');$caption=$block.kind -in @('figure_caption','table_caption')
    if($caption){$keepLines=Ensure-W (Ensure-W $p 'pPr') 'keepLines';Set-W $keepLines 'val' 1}
    $captionText=$block.text;$captionNote=$null
    if($caption){
        $captionNote=$block.note
        # paper_overrides.json: table_caption_paren_to_note = {tables:[n..], pattern:'^(.+?)(（…）)$'}
        $moveRule=$Overrides.table_caption_paren_to_note
        if($moveRule -and $block.kind -eq 'table_caption' -and $block.number -in @($moveRule.tables)){
            if($captionText -match [string]$moveRule.pattern){$captionText=$Matches[1].TrimEnd();$captionNote=($Matches[2]+' '+$captionNote).Trim()}
        }
    }
    $align=if($caption -or $block.kind -eq 'equation'){'center'}elseif($heading){'left'}else{'both'}
    Set-Paragraph $p $align $(if($block.kind -eq 'paragraph'){[int]([double]$ParaRules.body_first_line_chars*100)}else{0}) $(if($caption){$CaptionLineSpacing}elseif($block.kind -eq 'reference'){[double]$ParaRules.reference_line_spacing}elseif($block.kind -eq 'equation'){[double]$ParaRules.equation_line_spacing}else{[double]$ParaRules.line_spacing}) ($heading -or $block.kind -eq 'table_caption')
    if($block.kind -eq 'reference'){
        $ind=Ensure-W (Ensure-W $p 'pPr') 'ind';$ind.RemoveAttribute('firstLineChars',$W);$ind.RemoveAttribute('firstLine',$W);Set-W $ind 'left' $RefHang;Set-W $ind 'hanging' $RefHang
    }
    if($block.kind -eq 'references_heading'){$referencesStarted=$true}
    if($block.kind -eq 'heading' -and $block.text -match $MdRules.appendix_heading_pattern){$referencesStarted=$false}
    if($block.kind -eq 'equation'){Add-Formula $p $block.text $true $block.source_start_line}
    elseif($caption){Add-Inline $p $captionText $role $true $block.source_start_line}
    elseif($block.kind -eq 'paragraph' -and $block.text -match $MdRules.table_note_pattern -and $index -gt 0 -and $Blocks[$index-1].kind -eq 'table'){
        Set-Paragraph $p 'both' 0 $CaptionLineSpacing $false;Add-Inline $p $block.text $RoleNames.body $true $block.source_start_line ([double]$Rules.captions.note_size_pt)
    }
    else{Add-Inline $p $block.text $role ($block.kind -ne 'reference') $block.source_start_line}
    if($caption -and $captionNote -and $block.kind -eq 'figure_caption'){
        $keep=Ensure-W (Ensure-W $p 'pPr') 'keepNext';Set-W $keep 'val' 1
    }
    if($caption -and $block.kind -eq 'table_caption'){
        $group='tables';$englishRole=$RoleNames.table_caption_secondary
    }elseif($caption){$group='figures';$englishRole=$RoleNames.figure_caption_secondary}
    # Keep the caption, its secondary-language counterpart (english_captions.json), and the object together.
    if($caption -and $EnglishCaptions[$group].ContainsKey($block.number)){
        $keep=Ensure-W (Ensure-W $p 'pPr') 'keepNext';Set-W $keep 'val' 1
    }
    Append-P $p;Add-Ledger $block $p $role
    if($caption -and $EnglishCaptions[$group].ContainsKey($block.number)){
        $ep=New-Paragraph $englishRole;Set-Paragraph $ep 'center' 0 $CaptionLineSpacing ($block.kind -eq 'table_caption' -or [bool]$captionNote)
        $keepLines=Ensure-W (Ensure-W $ep 'pPr') 'keepLines';Set-W $keepLines 'val' 1
        $englishText=([string]$Rules.captions.secondary_format[$group]).Replace('{n}',[string]$block.number).Replace('{title}',[string]$EnglishCaptions[$group][$block.number])
        Add-TextRun $ep $englishText $englishRole;Append-P $ep
        $ContentLedger.Add([ordered]@{kind='english_caption';role=$englishRole;source_block_id=$block.source_block_id;source_start_line=0;source_end_line=0;source_text=$englishText;rendered_text=$englishText;source='approved English caption map';layout_relocation=$block.layout_relocation})
    }
    if($caption -and $captionNote){
        if($block.kind -eq 'figure_caption'){Add-CaptionNote $block $captionNote}
        else{$PendingTableNotes[$block.number]=@{block=$block;text=$captionNote}}
    }
}
# Closing sections (rules.sections.tail, in order; the last one becomes the body's final sectPr).
# rules.author_info_placeholder (optional) goes into the first closing section.
Add-SectionEnd $SectionTemplates[$SecBody] $script:CurrentColumns
$authorPlaceholder=[string]$Rules.author_info_placeholder
if($authorPlaceholder){
    $p=New-Paragraph $RoleNames.body;Set-Paragraph $p 'left' 0 1 $false;Add-TextRun $p $authorPlaceholder $RoleNames.body;Append-P $p
    $ContentLedger.Add([ordered]@{kind='generated_placeholder';role='author_information';source_start_line=0;source_end_line=0;source_text=$authorPlaceholder;rendered_text=$authorPlaceholder;reason='author biography, corresponding-author and funding facts are not supplied'})
}
for($t=0;$t -lt $SecTail.Count-1;$t++){Add-SectionEnd $SectionTemplates[$SecTail[$t]] ([Math]::Max(1,[int]$Mapping.sections[$SecTail[$t]].columns))}
$null=$Body.AppendChild($SectionTemplates[$SecTail[-1]].CloneNode($true))

# Template sample text in headers/footers (profile.json "header_fixes" / "footer_fixes"):
# runs are edited in place so PAGE fields and paragraph structure survive.
foreach($fix in @($JournalProfile.header_fixes)){
    if(-not $fix){continue}
    $partName=[string]$fix.part
    $header=if($Replacements.ContainsKey($partName)){Read-Xml $Replacements[$partName]}else{Read-Xml $Parts[$partName]}
    $headerNs=[Xml.XmlNamespaceManager]::new($header.NameTable);$headerNs.AddNamespace('w',$W)
    if($fix.replace_run_compact){
        # a whole run whose text (whitespace removed) equals the sample value
        $journalRuns=@($header.SelectNodes('//w:t',$headerNs)|Where-Object{$_.InnerText -replace '\s','' -eq [string]$fix.replace_run_compact})
        if($journalRuns.Count -ne 1){throw "Expected exactly one '$($fix.replace_run_compact)' run in $partName."}
        $journalRuns[0].InnerText=[string]$fix.with
    }elseif($fix.replace_span){
        # text split over several runs (WPS-converted templates): replace a character span of the match
        $headerTextNodes=@($header.SelectNodes('//w:t',$headerNs))
        $headerText=($headerTextNodes|ForEach-Object{$_.InnerText})-join ''
        $occurrences=[regex]::Matches($headerText,[regex]::Escape([string]$fix.replace_span))
        if($occurrences.Count -ne 1){throw "Expected one '$($fix.replace_span)' in $partName."}
        $numberStart=$occurrences[0].Index+[int]$fix.span_offset;$numberEnd=$numberStart+[int]$fix.span_length;$offset=0;$inserted=$false
        foreach($textNode in $headerTextNodes){
            $oldText=$textNode.InnerText;$nodeStart=$offset;$nodeEnd=$offset+$oldText.Length
            $overlapStart=[Math]::Max($nodeStart,$numberStart);$overlapEnd=[Math]::Min($nodeEnd,$numberEnd)
            if($overlapStart -lt $overlapEnd){
                $replacement=if(-not $inserted){[string]$fix.with}else{''}
                $textNode.InnerText=$oldText.Substring(0,$overlapStart-$nodeStart)+$replacement+$oldText.Substring($overlapEnd-$nodeStart)
                $inserted=$true
            }
            $offset=$nodeEnd
        }
        $patchedHeaderText=($headerTextNodes|ForEach-Object{$_.InnerText})-join ''
        if(-not $patchedHeaderText.Contains([string]$fix.expect) -or $patchedHeaderText.Contains([string]$fix.replace_span)){throw "Span replacement '$($fix.replace_span)' did not complete in $partName."}
    }
    $Replacements[$partName]=Xml-Bytes $header
}
foreach($fix in @($JournalProfile.footer_fixes)){
    if(-not $fix){continue}
    $partName=[string]$fix.part
    $footer=Read-Xml $Parts[$partName];$footerNs=[Xml.XmlNamespaceManager]::new($footer.NameTable);$footerNs.AddNamespace('w',$W)
    foreach($t in $footer.SelectNodes('//w:t',$footerNs)){
        foreach($pair in @($fix.replace)){if($pair){$t.InnerText=$t.InnerText.Replace([string]$pair[0],[string]$pair[1])}}
        foreach($pair in @($fix.exact)){if($pair -and $t.InnerText -eq [string]$pair[0]){$t.InnerText=[string]$pair[1]}}
        if($fix.append_suffix -and $t.InnerText -in @($fix.append_suffix.when)){$t.InnerText+=[string]$fix.append_suffix.suffix}
    }
    $Replacements[$partName]=Xml-Bytes $footer
}
$settings=Read-Xml $Parts['word/settings.xml'];$settingsNs=[Xml.XmlNamespaceManager]::new($settings.NameTable);$settingsNs.AddNamespace('w',$W)
$update=$settings.SelectSingleNode('/w:settings/w:updateFields',$settingsNs)
if(-not $update){$update=$settings.CreateElement('w','updateFields',$W);$null=$settings.DocumentElement.AppendChild($update)}
$null=$update.SetAttribute('val',$W,'true');$Replacements['word/settings.xml']=Xml-Bytes $settings
$core=Read-Xml $Parts['docProps/core.xml'];$coreNs=[Xml.XmlNamespaceManager]::new($core.NameTable)
$coreNs.AddNamespace('dc','http://purl.org/dc/elements/1.1/');$coreNs.AddNamespace('cp','http://schemas.openxmlformats.org/package/2006/metadata/core-properties')
foreach($node in $core.SelectNodes('//dc:creator|//cp:lastModifiedBy',$coreNs)){$node.InnerText=''}
$titleNode=$core.SelectSingleNode('//dc:title',$coreNs);if($titleNode){$titleNode.InnerText=$Front[[string]$FrontRules.document_title_role].text}
$Replacements['docProps/core.xml']=Xml-Bytes $core
$app=Read-Xml $Parts['docProps/app.xml']
foreach($node in @($app.DocumentElement.ChildNodes|Where-Object{$_.LocalName -in @('Pages','Words','Characters','CharactersWithSpaces','Paragraphs','Lines')})){$null=$node.ParentNode.RemoveChild($node)}
foreach($node in @($app.DocumentElement.ChildNodes|Where-Object{$_.LocalName -eq 'Company'})){$node.InnerText=''}
$Replacements['docProps/app.xml']=Xml-Bytes $app
# The original sample figures/portraits must not remain hidden in the new package.
# Relationship IDs are local to their source part, so inspect document.xml usage
# first, then retain any candidate media still targeted by another relationship.
$usedDocumentRelationships=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach($attribute in $Doc.SelectNodes("//@*[namespace-uri()='$R' or (local-name()='relid' and namespace-uri()='urn:schemas-microsoft-com:office:office')]")){$null=$usedDocumentRelationships.Add($attribute.Value)}
$orphanCandidates=[Collections.Generic.List[string]]::new()
foreach($relationship in @($DocumentRels.DocumentElement.ChildNodes|Where-Object{$_.NodeType -eq 'Element' -and $_.GetAttribute('Type') -eq ($R+'/image')})){
    $id=$relationship.GetAttribute('Id')
    if(-not $usedDocumentRelationships.Contains($id)){
        $target=$relationship.GetAttribute('Target')
        if($relationship.GetAttribute('TargetMode') -eq 'External'){throw 'Unexpected external sample-image relationship; review before removal.'}
        $resolved=[Uri]::UnescapeDataString([Uri]::new([Uri]::new('https://package.invalid/word/document.xml'),$target).AbsolutePath.TrimStart('/'))
        if($resolved -notmatch '^word/media/'){throw "Unexpected unreferenced image target: $resolved"}
        $orphanCandidates.Add($resolved)
        $RemovedRelationships.Add([ordered]@{part='word/_rels/document.xml.rels';id=$id;target=$target;resolved_part=$resolved;reason='original template sample image not referenced by the generated document'})
        $null=$DocumentRels.DocumentElement.RemoveChild($relationship)
    }
}
$remainingRelationshipTargets=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
foreach($relationshipPart in @($Parts.Keys|Where-Object{$_ -match '\.rels$'})){
    $relationships=if($relationshipPart -eq 'word/_rels/document.xml.rels'){$DocumentRels}else{Read-Xml $Parts[$relationshipPart]}
    $sourcePart=if($relationshipPart -eq '_rels/.rels'){''}else{$relationshipPart -replace '/_rels/([^/]+)\.rels$','/$1'}
    foreach($relationship in $relationships.DocumentElement.ChildNodes|Where-Object{$_.NodeType -eq 'Element' -and $_.GetAttribute('TargetMode') -ne 'External'}){
        $resolved=[Uri]::UnescapeDataString([Uri]::new([Uri]::new('https://package.invalid/'+$sourcePart),$relationship.GetAttribute('Target')).AbsolutePath.TrimStart('/'))
        $null=$remainingRelationshipTargets.Add($resolved)
    }
}
foreach($candidate in $orphanCandidates){if(-not $remainingRelationshipTargets.Contains($candidate)){$null=$RemovedParts.Add($candidate)}}
# Keep generated property children in WordprocessingML schema order.  Source
# content and package preservation remain separate from this local normalization.
$orders=@{
    pPr='pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'
    rPr='rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange'
    tblPr='tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription tblPrChange'
    trPr='cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden ins del trPrChange'
    tcPr='cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark headers cellIns cellDel cellMerge tcPrChange'
    sectPr='headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange'
}
foreach($propertyName in $orders.Keys){
    $rank=@{};$ordinal=0;foreach($name in $orders[$propertyName].Split(' ')){$rank[$name]=$ordinal;$ordinal++}
    foreach($propertyNode in $Doc.SelectNodes("//w:$propertyName",$Ns)){
        $children=@($propertyNode.ChildNodes|Where-Object{$_.NodeType -eq 'Element'}|Sort-Object -Stable {if($rank.ContainsKey($_.LocalName)){$rank[$_.LocalName]}else{1000}})
        foreach($child in $children){$null=$propertyNode.AppendChild($child)}
    }
}
$Replacements['word/document.xml']=Xml-Bytes $Doc
$Replacements['word/_rels/document.xml.rels']=Xml-Bytes $DocumentRels
$Replacements['[Content_Types].xml']=Xml-Bytes $ContentTypes

$bodyText=($Body.SelectNodes('.//w:t|.//m:t',$Ns)|ForEach-Object{$_.InnerText})-join ''
$ledgerText=($ContentLedger|ForEach-Object{$_.rendered_text})-join ''
if($bodyText -cne $ledgerText){throw 'The complete body text is not covered exactly by the content ledger.'}
$counts=[ordered]@{headings=@($Blocks|Where-Object{$_.kind -in @('heading','references_heading')}).Count+1;images=$ImageManifest.Count;tables=$tableNumber;display_formulas=@($FormulaManifest|Where-Object{$_.display}).Count;inline_formulas=@($FormulaManifest|Where-Object{-not $_.display}).Count;references=@($Blocks|Where-Object{$_.kind -eq 'reference'}).Count;source_nonblank_lines=$Coverage.Count;source_blocks=$Blocks.Count}
$sourceText=[IO.File]::ReadAllText($SourceFile,[Text.Encoding]::UTF8)
$withoutDisplay=[regex]::Replace($sourceText,'(?s)\$\$.*?\$\$','')
$expectedCounts=[ordered]@{
    headings=[regex]::Matches($sourceText,'(?m)^#{1,6}\s+').Count
    images=[regex]::Matches($sourceText,'!\[[^\]]*\]\([^)]+\)').Count
    tables=[regex]::Matches($sourceText,'(?m)^\|[^\r\n]*\r?\n(?:\|[^\r\n]*(?:\r?\n|$))+').Count
    display_formulas=[regex]::Matches($sourceText,'(?s)\$\$(.*?)\$\$').Count
    inline_formulas=[regex]::Matches($withoutDisplay,'\$[^$\r\n]+\$').Count
    references=[regex]::Matches($sourceText,'(?m)^\[\d+\]\s+').Count
}
foreach($key in $expectedCounts.Keys){if($counts[$key] -ne $expectedCounts[$key]){throw "Source/object count mismatch for $key : source $($expectedCounts[$key]), constructed $($counts[$key])"}}
if($Doc.SelectNodes('//m:oMath',$Ns).Count -ne $FormulaManifest.Count){throw 'Formula object count does not match the source manifest.'}
if($Doc.SelectNodes('//w:tbl',$Ns).Count -ne $tableNumber){throw 'Table object count does not match the source inventory.'}
if($Doc.SelectNodes('//w:drawing',$Ns).Count -ne $ImageManifest.Count){throw 'Image object count does not match the source inventory.'}
$sectionRecords=@($Doc.SelectNodes('//w:sectPr',$Ns)|ForEach-Object{[ordered]@{columns=$_.SelectSingleNode('w:cols',$Ns).GetAttribute('num',$W);xml=$_.OuterXml}})
$preserved=@($Parts.Keys|Where-Object{-not $Replacements.ContainsKey($_) -and -not $RemovedParts.Contains($_)})
$report=[ordered]@{
    status='constructed_pending_word_render_and_mathtype_review';authoring_gate='native_word_visual_verification_pending'
    source_path=$SourceFile;source_sha256=$SourceHash;reference_path=$ReferenceFile;reference_sha256=$ReferenceHash;output_path=$OutputFile;report_directory=$ReportDirectoryPath
    layout_config_path=[IO.Path]::GetFullPath($LayoutConfigPath);figure_map_path=[IO.Path]::GetFullPath($FigureMapPath)
    counts=$counts;modified_parts=@($Replacements.Keys|Sort-Object);added_parts=@($AddedParts.Keys|Sort-Object);removed_parts=@($RemovedParts|Sort-Object);removed_relationships=$RemovedRelationships;preserved_part_count=$preserved.Count
    content_ledger_complete=$true;content_ledger_entries=$ContentLedger.Count;body_text_characters=$bodyText.Length;ledger_text_characters=$ledgerText.Length
    preserved_parts=@($preserved|Sort-Object|ForEach-Object{[ordered]@{part=$_;sha256=$PartHashes[$_]}})
    sections=$sectionRecords;figures=$ImageManifest;full_width_tables=$FullWidthTables;full_width_figures=$FullWidthFigures;full_width_equations=$FullWidthEquations
    caption_line_spacing_multiplier=$CaptionLineSpacing;caption_line_spacing_twips=[int][Math]::Round($CaptionLineSpacing*240)
    table_layouts=$TableLayoutManifest;layout_relocations=$LayoutRelocationManifest;scientific_prose_order_unchanged=$true;relative_table_order_unchanged=$true;prose_visual_interleaving_changed_only_as_configured=$true
    source_block_manifest=@(for($i=0;$i -lt $Blocks.Count;$i++){[ordered]@{id=$Blocks[$i].source_block_id;kind=$Blocks[$i].kind;is_table_note=[bool]$Blocks[$i].is_table_note;source_start_line=$Blocks[$i].source_start_line;source_end_line=$Blocks[$i].source_end_line;original_index=$Blocks[$i].original_block_index;rendered_index=$i;layout_relocation=$Blocks[$i].layout_relocation}})
    english_caption_count=@($ContentLedger|Where-Object{$_.kind -eq 'english_caption'}).Count
    reference_entry_count=$ReferenceManifest.Count;bilingual_reference_count=@($ReferenceManifest|Where-Object{$_.chinese}).Count
    references=$ReferenceManifest
    intentional_template_changes=@('replace the template sample body with the Markdown content','clone the template''s own sections (continuous breaks) for full-width objects; first-page settings removed on added sections','apply profile.json header_fixes/footer_fixes in place ('+@(@($JournalProfile.header_fixes)+@($JournalProfile.footer_fixes)|Where-Object{$_}).Count+' rule(s)); keep fields','remove stale page/word-count caches, author/company metadata and unreferenced template media','keep short tables, their captions and secondary captions together; caption notes as separate paragraphs','apply references_formatted.json display data when supplied; original entries kept in the reference manifest','optional generated lines from the profile (classification line, author-information placeholder) and paper overrides (title line break)')
    limitations=@('not visually verified; the host (WPS/Word) must open, update fields, and render all pages (finalize_native.ps1, export_pdf.ps1)','OMML is editable intermediate math; MathType compliance remains unverified','font and print-size legibility require visual review','scientific prose is imported without editorial changes (see source_numbering_correction when supplied); bibliographic display transformations are recorded separately')
    formula_manifest='formula_manifest.json';content_ledger='manuscript_content_ledger.json'
}
if(Test-Path -LiteralPath $NumberingCorrectionPath){$report['source_numbering_correction']=Get-Content -LiteralPath $NumberingCorrectionPath -Raw|ConvertFrom-Json -Depth 8}
$report['math_engine']=$MathEngine;$report['journal_profile']=$ProfileDir
if((Get-FileHash -Algorithm SHA256 -LiteralPath $SourceFile).Hash -ne $SourceHash){throw 'Manuscript changed during the build.'}
if((Get-FileHash -Algorithm SHA256 -LiteralPath $ReferenceFile).Hash -ne $ReferenceHash){throw 'Protected reference changed during the build.'}
if($ValidateOnly){$report|ConvertTo-Json -Depth 15;return}

if(-not(Test-Path -LiteralPath $OutputDirectory)){throw 'Output directory must already exist.'}
if(-not(Test-Path -LiteralPath $ReportDirectoryPath)){throw 'Report directory must already exist.'}
$staging=Join-Path $OutputDirectory ('.manuscript_working.'+[Guid]::NewGuid().ToString('N')+'.docx')
try{
    $sourceZip=[IO.Compression.ZipFile]::OpenRead($ReferenceFile)
    $outputStream=[IO.File]::Open($staging,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
    $outputZip=[IO.Compression.ZipArchive]::new($outputStream,[IO.Compression.ZipArchiveMode]::Create,$false)
    try{
        foreach($entry in $sourceZip.Entries){
            if($RemovedParts.Contains($entry.FullName)){continue}
            $out=$outputZip.CreateEntry($entry.FullName,[IO.Compression.CompressionLevel]::Optimal);$out.LastWriteTime=$entry.LastWriteTime
            # An if-expression enumerates an empty byte array into $null in PowerShell.
            # Assign inside the branch so legitimate zero-length ZIP entries survive.
            if($Replacements.ContainsKey($entry.FullName)){$bytes=$Replacements[$entry.FullName]}else{$bytes=$Parts[$entry.FullName]}
            if($null -eq $bytes){throw "Package bytes unexpectedly null: $($entry.FullName)"}
            $stream=$out.Open();try{if($bytes.Length -gt 0){$stream.Write($bytes,0,$bytes.Length)}}finally{$stream.Dispose()}
        }
        foreach($part in $AddedParts.Keys|Sort-Object){if($null -eq $AddedParts[$part]){throw "Added package bytes unexpectedly null: $part"};$out=$outputZip.CreateEntry($part,[IO.Compression.CompressionLevel]::Optimal);$stream=$out.Open();try{if($AddedParts[$part].Length -gt 0){$stream.Write($AddedParts[$part],0,$AddedParts[$part].Length)}}finally{$stream.Dispose()}}
    }finally{$outputZip.Dispose();$outputStream.Dispose();$sourceZip.Dispose()}
    $checkZip=[IO.Compression.ZipFile]::OpenRead($staging)
    try{
        foreach($part in $preserved){$entry=$checkZip.GetEntry($part);if(-not $entry){throw "Preserved part missing: $part"};$stream=$entry.Open();$memory=[IO.MemoryStream]::new();try{$stream.CopyTo($memory);if((Hash-Bytes $memory.ToArray()) -ne $PartHashes[$part]){throw "Preserved part changed: $part"}}finally{$memory.Dispose();$stream.Dispose()}}
    }finally{$checkZip.Dispose()}
    Move-Item -LiteralPath $staging -Destination $OutputFile -Force:$Force
}finally{if(Test-Path -LiteralPath $staging){Remove-Item -LiteralPath $staging}}
$report['output_sha256']=(Get-FileHash -Algorithm SHA256 -LiteralPath $OutputFile).Hash
[IO.File]::WriteAllText((Join-Path $ReportDirectoryPath 'formula_manifest.json'),($FormulaManifest|ConvertTo-Json -Depth 8),$Utf8)
[IO.File]::WriteAllText((Join-Path $ReportDirectoryPath 'manuscript_content_ledger.json'),($ContentLedger|ConvertTo-Json -Depth 8),$Utf8)
[IO.File]::WriteAllText((Join-Path $ReportDirectoryPath 'manuscript_build_report.json'),($report|ConvertTo-Json -Depth 15),$Utf8)
$report|ConvertTo-Json -Depth 15
