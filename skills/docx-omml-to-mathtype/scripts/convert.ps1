<#
One-shot: convert every Word-native (OMML) equation in a .docx into a MathType (Equation.DSMT4)
OLE object, using MathType's API + WPS (or Word) COM. No VBA needed.

  pwsh -File convert.ps1 -InputDocx paper.docx [-OutputDocx paper_mathtype.docx] [-FontSizePt 10.5]
       [-CjkFont SimSun] [-Only 1,7] [-ComProgId KWps.Application] [-Xsl OMML2MML.XSL] [-Force] [-SkipPdf]

Writes the result to -OutputDocx (default: <input>_mathtype.docx), never touches the input.
Work files (marked docx, MathML json, worker logs, GIFs, PDF, page PNGs) go to -WorkDir.
Exit code 0 = all checks passed.
#>
param(
    [Parameter(Mandatory)] [string]$InputDocx,
    [string]$OutputDocx,
    [double]$FontSizePt = 0,              # 0 = body size from the document's Normal style
    [string]$CjkFont = "SimSun",          # MathType "Text" style font, used for CJK inside equations
    [string]$Only,                        # e.g. "1,7" to convert a subset (pilot)
    [string]$ComProgId = "KWps.Application",   # "Word.Application" for MS Word (untested)
    [string]$Xsl,                         # OMML2MML.XSL if not in a standard Office folder
    [string]$WorkDir,
    [switch]$Force,
    [switch]$SkipPdf
)
$ErrorActionPreference = "Stop"
$s = $PSScriptRoot
$InputDocx = (Resolve-Path $InputDocx).Path
if (-not $OutputDocx) { $OutputDocx = [IO.Path]::ChangeExtension($InputDocx, $null).TrimEnd('.') + "_mathtype.docx" }
$OutputDocx = [IO.Path]::GetFullPath($OutputDocx)
if ($OutputDocx -eq $InputDocx) { throw "OutputDocx must differ from InputDocx (keep the original as backup)" }
if ((Test-Path $OutputDocx) -and -not $Force) { throw "$OutputDocx exists; pass -Force to overwrite" }
if (-not $WorkDir) { $WorkDir = Join-Path $env:TEMP ("docx-omml-to-mathtype\" + (Get-Date -Format "yyyyMMdd-HHmmss")) }
New-Item -ItemType Directory -Force $WorkDir | Out-Null
"work dir: $WorkDir"

# ---------- preflight (see references/requirements.md) ----------
$checkArgs = @("-NoProfile", "-File", "$s\check_runtime.ps1", "-ComProgId", $ComProgId, "-CjkFont", $CjkFont)
if ($Xsl) { $checkArgs += @("-Xsl", $Xsl) }
pwsh @checkArgs
if ($LASTEXITCODE -ne 0) { throw "runtime requirements missing (see above)" }
$eqpBase = "C:\Program Files (x86)\MathType\Preferences\Times+Symbol 12.eqp"
$dir = Split-Path $InputDocx; $leaf = Split-Path $InputDocx -Leaf
foreach ($lock in @("~`$" + $leaf, "~`$" + $leaf.Substring([Math]::Min(2, $leaf.Length - 1)))) {
    if (Test-Path (Join-Path $dir $lock)) { Write-Warning "lock file $lock found: the document may be open in WPS/Word (fine, we only read it)" }
}

# ---------- python env (lxml for XSLT, pymupdf for page renders) ----------
$venv = Join-Path $env:LOCALAPPDATA "docx-omml-to-mathtype\venv"
$py = Join-Path $venv "Scripts\python.exe"
$envOk = $false
if (Test-Path $py) {
    & $py -I -c "import lxml, pymupdf" 2>$null
    $envOk = ($LASTEXITCODE -eq 0)
}
if (-not $envOk) {
    "setting up python env at $venv"
    if (Get-Command uv -ErrorAction SilentlyContinue) {
        uv venv -q --allow-existing --python 3.12 $venv; if (-not $?) { throw "uv venv failed" }
        uv pip install -q --python $py lxml pymupdf; if (-not $?) { throw "uv pip install failed" }
    } else {
        py -3 -m venv $venv; & $py -m pip install -q lxml pymupdf
    }
}

# ---------- MathType preferences: body font size + CJK text font ----------
if ($FontSizePt -le 0) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($InputDocx)
    try {
        $styles = (New-Object IO.StreamReader ($zip.GetEntry("word/styles.xml").Open())).ReadToEnd()
    } finally { $zip.Dispose() }
    $normal = [regex]::Match($styles, '<w:style w:type="paragraph" w:default="1".*?</w:style>', 'Singleline').Value
    $sz = [regex]::Match($normal, '<w:sz w:val="(\d+)"').Groups[1].Value
    if (-not $sz) { $sz = [regex]::Match($styles, '<w:rPrDefault>.*?<w:sz w:val="(\d+)"', 'Singleline').Groups[1].Value }
    $FontSizePt = if ($sz) { [int]$sz / 2 } else { 10.5 }
}
"equation size: $FontSizePt pt, CJK text font: $CjkFont"
$eqp = Join-Path $WorkDir "prefs.eqp"
(Get-Content $eqpBase -Raw) -replace '(?m)^Full=.*$', "Full=$FontSizePt pt" -replace '(?m)^Text=.*$', "Text=$CjkFont" |
    Set-Content $eqp -Encoding ascii -NoNewline

# ---------- 1. OMML -> markers + MathML ----------
$marked = Join-Path $WorkDir "marked.docx"
$eqJson = Join-Path $WorkDir "eqs.json"
$prepArgs = @("-I", "$s\prep_docx.py", $InputDocx, $marked, $eqJson)
if ($Only) { $prepArgs += @("--only", $Only) }
if ($Xsl) { $prepArgs += @("--xsl", $Xsl) }
& $py @prepArgs; if (-not $?) { throw "prep_docx.py failed" }

# ---------- 2. markers -> MathType objects (WPS/Word COM) ----------
$converted = Join-Path $WorkDir "converted.docx"
Copy-Item $marked $converted -Force
pwsh -NoProfile -STA -File "$s\wps_replace.ps1" -DocPath $converted -EqJson $eqJson -PrefsFile $eqp -WorkDir $WorkDir -ComProgId $ComProgId
$replaceOk = ($LASTEXITCODE -eq 0)

# ---------- 3. restore bookmarks the paste deleted ----------
$final = Join-Path $WorkDir "final.docx"
if (-not $Only) {
    & $py -I "$s\restore_bookmarks.py" $InputDocx $converted $final
    if (-not $?) { Write-Warning "bookmark restore skipped"; Copy-Item $converted $final -Force }
} else { Copy-Item $converted $final -Force }

# ---------- 4. verify ----------
"--- verification ---"
& $py -I "$s\verify_docx.py" $InputDocx $final
$verifyOk = $?

# ---------- 5. PDF + page renders for visual review ----------
if (-not $SkipPdf) {
    $pdf = Join-Path $WorkDir "final.pdf"
    pwsh -NoProfile -STA -File "$s\export_pdf.ps1" -DocPath $final -PdfPath $pdf -ComProgId $ComProgId
    if (Test-Path $pdf) {
        $png = Join-Path $WorkDir "pages"; New-Item -ItemType Directory -Force $png | Out-Null
        $n = & $py -I -c "import pymupdf,sys; print(pymupdf.open(sys.argv[1]).page_count)" $pdf
        & $py -I "$s\render_pages.py" $pdf $png --pages ((1..[int]$n) -join ",") | Select-Object -First 1
        "page images: $png"
    }
}

Copy-Item $final $OutputDocx -Force
"output: $OutputDocx"
if ($Only -and $replaceOk) { "PARTIAL RUN (-Only $Only): equations not selected stay OMML, so 'no OMML equations left' and 'one MathType object per equation' are expected to fail"; exit 0 }
if (-not ($replaceOk -and $verifyOk)) { "SOME CHECKS FAILED - see above"; exit 1 }
"ALL CHECKS PASSED"
