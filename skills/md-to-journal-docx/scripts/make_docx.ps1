<#
One-shot: Markdown manuscript [+ journal template] -> native-saved DOCX (+ DOC copy) with Word-native
equations, page renders for review, and evidence reports. MathType conversion is NOT part of this
(use the docx-omml-to-mathtype skill on the result).

Two inputs decide the mode:
  no template   -> plain: bundled styles, assets/plain_reference[_zh].docx (plain_docx.py)
  -TemplatePath -> template: the profile made for that template file (found by SHA-256);
                   none yet -> evidence for making one is prepared, exit 2 (resolve_template.ps1)
  (-ProfileName / -ProfileDir name a profile directly.)

  pwsh -File make_docx.ps1 -ManuscriptPath paper.md [-TemplatePath journal.doc] -OutputDirectory out -OutputStem paper_v2
       [-PaperConfigDir cfg] [-MathEngine pandoc|cs] [-ComProgId auto|KWps.Application|Word.Application]
       [-WorkDir dir] [-SkipHost] [-Force] [-AllowLiteralDollar]
       [-FigureMapPath …] [-LayoutConfigPath …] [-ReferenceFormattingPath …] [-EnglishCaptionMapPath …]
       [-ReferenceTranslationsPath …] [-PaperOverridesPath …] [-NumberingCorrectionPath …]   (template mode)

Steps: check_runtime -> build (validate, then write) -> host probe -> native save via host (fields,
DOCX+DOC, reopen, metrics) -> PDF + PNG of every page. Exit 0 = all automated checks passed; the
page images still need a human/agent look. Exit 2 = template has no profile yet.
#>
param(
    [Parameter(Mandatory)] [string]$ManuscriptPath,
    [string]$TemplatePath = '',
    [string]$ProfileName = '',
    [string]$ProfileDir = '',
    [string]$PaperConfigDir = '',
    [Parameter(Mandatory)] [string]$OutputDirectory,
    [Parameter(Mandatory)] [string]$OutputStem,
    [ValidateSet('pandoc','cs')] [string]$MathEngine = 'pandoc',
    [string]$ComProgId = 'auto',
    [string]$WorkDir = '',
    [switch]$SkipHost,
    [switch]$Force,
    [switch]$AllowLiteralDollar,
    [string]$FigureMapPath = '', [string]$LayoutConfigPath = '', [string]$ReferenceFormattingPath = '',
    [string]$EnglishCaptionMapPath = '', [string]$ReferenceTranslationsPath = '', [string]$PaperOverridesPath = '',
    [string]$NumberingCorrectionPath = ''
)
$ErrorActionPreference = 'Stop'
$s = $PSScriptRoot
$ManuscriptPath = (Resolve-Path $ManuscriptPath).Path
$OutputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
if (-not $WorkDir) { $WorkDir = Join-Path $env:TEMP ("md-to-journal-docx\" + (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$WorkDir = [IO.Path]::GetFullPath($WorkDir)
New-Item -ItemType Directory -Force $WorkDir, $OutputDirectory | Out-Null
"work dir: $WorkDir"
$paperCfg = @($PaperConfigDir, $FigureMapPath, $LayoutConfigPath, $ReferenceFormattingPath, $EnglishCaptionMapPath,
              $ReferenceTranslationsPath, $PaperOverridesPath, $NumberingCorrectionPath) | Where-Object { $_ }

# ---------- mode ----------
if ($ProfileName -and -not $ProfileDir) { $ProfileDir = Join-Path (Split-Path $s) "profiles\$ProfileName" }
if (-not $ProfileDir -and $TemplatePath) {
    $hostForConvert = if ($ComProgId -ne 'auto') { $ComProgId } elseif (Test-Path 'Registry::HKEY_CLASSES_ROOT\KWps.Application') { 'KWps.Application' } else { 'Word.Application' }
    $res = pwsh -NoProfile -STA -File "$s\resolve_template.ps1" -TemplatePath $TemplatePath -WorkDir $WorkDir -ComProgId $hostForConvert
    $code = $LASTEXITCODE
    $res
    if ($code -eq 2) { exit 2 }
    if ($code -ne 0) { throw 'resolve_template failed (see above)' }
    $ProfileDir = (($res | Where-Object { $_ -like 'PROFILE=*' }) -replace '^PROFILE=', '')
}
$Plain = -not $ProfileDir
if ($Plain) {
    "mode: plain (no template) - bundled styles"
    if ($paperCfg) { throw 'paper config (-PaperConfigDir / -…Path) applies to template mode only' }
    if ($MathEngine -ne 'pandoc') { throw '-MathEngine cs applies to template mode only' }
} else {
    $ProfileDir = (Resolve-Path $ProfileDir).Path
    "mode: template - profile $ProfileDir"
}

# ---------- 0. runtime ----------
$checkArgs = @('-NoProfile', '-File', "$s\check_runtime.ps1", '-MathEngine', $MathEngine, '-ComProgId', $ComProgId)
if (-not $Plain) { $checkArgs += @('-ProfileDir', $ProfileDir) }
if ($SkipHost) { $checkArgs += '-SkipHost' }
pwsh @checkArgs
if ($LASTEXITCODE -ne 0) { throw 'runtime requirements missing (see above)' }
$venv = Join-Path $env:LOCALAPPDATA 'md-to-journal-docx\venv'; $py = Join-Path $venv 'Scripts\python.exe'
$envOk = $false
if (Test-Path $py) { & $py -I -c "import pypandoc, lxml, pymupdf" 2>$null; $envOk = ($LASTEXITCODE -eq 0) }
if (-not $envOk -and ($MathEngine -eq 'pandoc' -or -not $SkipHost)) {
    "setting up python env at $venv (pypandoc_binary, lxml, pymupdf)"
    if (Get-Command uv -ErrorAction SilentlyContinue) {
        uv venv -q --allow-existing --python 3.12 $venv; if (-not $?) { throw 'uv venv failed' }
        uv pip install -q --python $py pypandoc_binary lxml pymupdf; if (-not $?) { throw 'uv pip install failed' }
    } else { py -3 -m venv $venv; & $py -m pip install -q pypandoc_binary lxml pymupdf }
}

# ---------- 1. build ----------
$working = Join-Path $WorkDir 'manuscript_working.docx'
if ($Plain) {
    "--- build (pandoc, bundled styles)"
    $pa = @('-I', "$s\plain_docx.py", $ManuscriptPath, $working, $WorkDir)
    if ($AllowLiteralDollar) { $pa += '--allow-literal-dollar' }
    & $py @pa
    if ($LASTEXITCODE -ne 0) { throw 'plain build failed (errors above; fix the Markdown)' }
} else {
    $build = @{ ManuscriptPath = $ManuscriptPath; ProfileDir = $ProfileDir; MathEngine = $MathEngine; PythonPath = $py
                OutputPath = $working; ReportDirectory = $WorkDir }
    if ($PaperConfigDir) { $build.PaperConfigDir = (Resolve-Path $PaperConfigDir).Path }
    foreach ($n in 'FigureMapPath','LayoutConfigPath','ReferenceFormattingPath','EnglishCaptionMapPath','ReferenceTranslationsPath','PaperOverridesPath','NumberingCorrectionPath') {
        $v = Get-Variable $n -ValueOnly; if ($v) { $build[$n] = (Resolve-Path $v).Path }
    }
    "--- build (validate only)"
    & "$s\build_docx.ps1" @build -ValidateOnly | Out-Null
    "--- build"
    & "$s\build_docx.ps1" @build -Force | Out-Null
}
$report = Get-Content (Join-Path $WorkDir 'manuscript_build_report.json') -Raw -Encoding utf8 | ConvertFrom-Json
"built: " + ($report.counts | ConvertTo-Json -Compress) + "  ledger complete: $($report.content_ledger_complete)  engine: $($report.math_engine)"
if ($SkipHost) {
    Copy-Item $working (Join-Path $OutputDirectory "$OutputStem.docx") -Force:$Force
    "SkipHost: wrote the unsaved builder output (fields not updated, not page-checked): $(Join-Path $OutputDirectory "$OutputStem.docx")"
    exit 0
}

# ---------- 2a. host capability: prove it on a copy of this document, don't assume it ----------
$candidates = if ($ComProgId -ne 'auto') { @($ComProgId) } else { @('KWps.Application', 'Word.Application') | Where-Object { Test-Path "Registry::HKEY_CLASSES_ROOT\$_" } }
$chosen = $null
foreach ($c in $candidates) {
    $probeJson = Join-Path $WorkDir ("host_probe_" + $c.Split('.')[0] + ".json")
    pwsh -NoProfile -STA -File "$s\probe_host.ps1" -DocPath $working -ComProgId $c -OutJson $probeJson | Out-Null
    $pr = Get-Content $probeJson -Raw | ConvertFrom-Json
    "host probe $c : passed=$($pr.passed) ($($pr.version)) $(if (-not $pr.passed) { 'errors: ' + ($pr.errors -join '; ') })"
    if ($pr.passed) { $chosen = $c; break }
}
if (-not $chosen) { throw "no document host passed the capability probe ($($candidates -join ', ')); see host_probe_*.json in $WorkDir" }
$ComProgId = $chosen

# ---------- 2b. native save through the host ----------
"--- native save via $ComProgId"
$fin = @('-NoProfile', '-STA', '-File', "$s\finalize_native.ps1", '-SourcePath', $working,
         '-OutputDirectory', $OutputDirectory, '-OutputStem', $OutputStem, '-ComProgId', $ComProgId)
if ($Force) { $fin += '-Force' }
pwsh @fin | Out-Null
$check = Get-Content (Join-Path $OutputDirectory "${OutputStem}_save_check.json") -Raw -Encoding utf8 | ConvertFrom-Json
if ($check.status -ne 'passed') { throw "native save failed: $($check.error.message)" }
$d = $check.docx_check
"saved: pages=$($d.page_count) sections=$($d.section_count) tables=$($d.table_count) (spanning pages: $(@($d.tables | Where-Object spans_pages).Count)) omath=$($d.omath_count) inline_shapes=$($d.inline_shape_count)"
"doc copy: pages=$($check.doc_check.page_count) omath=$($check.doc_check.omath_count) (DOC turns equations into pictures)"
$problems = @()
if ($d.omath_count -ne ($report.counts.display_formulas + $report.counts.inline_formulas)) { $problems += "equation count $($d.omath_count) != source $($report.counts.display_formulas + $report.counts.inline_formulas)" }
if ($d.table_count -ne $report.counts.tables) { $problems += "table count $($d.table_count) != source $($report.counts.tables)" }
$spanning = (@($d.tables | Where-Object spans_pages) | ForEach-Object index) -join ','
if ($spanning -and $Plain) { "WARN tables spanning pages: $spanning (kept together where they fit; one taller than a page must break - check the pages)" }
elseif ($spanning) { $problems += "tables spanning pages: $spanning" }

# ---------- 3. renders ----------
$pdf = Join-Path $WorkDir "$OutputStem.pdf"
pwsh -NoProfile -STA -File "$s\export_pdf.ps1" -DocPath (Join-Path $OutputDirectory "$OutputStem.docx") -PdfPath $pdf -ComProgId $ComProgId | Out-Null
$pages = Join-Path $WorkDir 'pages'
& $py -I "$s\render_pages.py" $pdf $pages --pages all | Out-Null
"page images: $pages"
"outputs: $(Join-Path $OutputDirectory "$OutputStem.docx"), .doc, ${OutputStem}_save_check.json; build reports in $WorkDir"
if ($problems) { "CHECKS FAILED:"; $problems | ForEach-Object { "  - $_" }; exit 1 }
"ALL AUTOMATED CHECKS PASSED - now review the page images"
