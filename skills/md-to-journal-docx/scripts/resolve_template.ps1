<#
Template given -> which profile? Matches the template file's SHA-256 against every profile under
profiles/ (format_mapping.reference_sha256 = the DOCX the profile was made from,
template_visual_review.original_doc_sha256 = the journal's original file, e.g. a .doc).

  pwsh -File resolve_template.ps1 -TemplatePath journal.doc[x] -WorkDir <dir> [-ComProgId KWps.Application] [-ProfilesRoot <dir>]

Exit 0: profile found (last line: PROFILE=<dir>).
Exit 2: no profile. The facts for making one are prepared in <WorkDir>/template/:
        template.docx (a .doc/.dot/.dotx/.wps is converted through the host), template_probe.json/.md,
        template.pdf + pages/pageNN.png. Next: references/new_journal_profile.md.
#>
param([Parameter(Mandatory)] [string]$TemplatePath, [Parameter(Mandatory)] [string]$WorkDir,
      [string]$ComProgId = 'KWps.Application', [string]$ProfilesRoot = '')
$ErrorActionPreference = 'Stop'
$s = $PSScriptRoot
$TemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path
if (-not $ProfilesRoot) { $ProfilesRoot = Join-Path (Split-Path $s) 'profiles' }
$hash = (Get-FileHash -LiteralPath $TemplatePath -Algorithm SHA256).Hash
"template: $TemplatePath"
"sha256:   $hash"

foreach ($dir in Get-ChildItem -LiteralPath $ProfilesRoot -Directory -ErrorAction SilentlyContinue) {
    $known = @()
    $map = Join-Path $dir.FullName 'format_mapping.json'
    $rev = Join-Path $dir.FullName 'template_visual_review.json'
    if (Test-Path $map) { $known += (Get-Content $map -Raw -Encoding utf8 | ConvertFrom-Json).reference_sha256 }
    if (Test-Path $rev) { $known += (Get-Content $rev -Raw -Encoding utf8 | ConvertFrom-Json).original_doc_sha256 }
    if ($known -contains $hash) { "PROFILE=$($dir.FullName)"; exit 0 }
}

# ---- no profile: prepare the evidence for making one ----
$out = Join-Path ([IO.Path]::GetFullPath($WorkDir)) 'template'
New-Item -ItemType Directory -Force $out | Out-Null
$docx = Join-Path $out 'template.docx'
if ([IO.Path]::GetExtension($TemplatePath) -ieq '.docx') { Copy-Item -LiteralPath $TemplatePath $docx -Force }
else {
    # .doc/.dot/.dotx/.wps: the probe reads OOXML, so save a DOCX copy through the host (original untouched)
    $before = @(Get-Process wps, winword -ErrorAction SilentlyContinue | ForEach-Object Id)
    $app = New-Object -ComObject $ComProgId
    $foreign = $app.Documents.Count
    try {
        if ($foreign -ne 0) { throw "$ComProgId instance already holds documents (in use by the user) - not touched; close them or convert the template to .docx by hand" }
        $app.Visible = $false; try { $app.DisplayAlerts = 0 } catch {}
        $d = $app.Documents.Open($TemplatePath, $false, $true)
        try { $d.SaveAs2($docx, 16) } catch { $d.SaveAs($docx, 16) }
        $d.Close(0)
    } finally {
        if ($foreign -eq 0) { try { $app.Quit() } catch {} }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
        Start-Sleep -Seconds 2
        Get-CimInstance Win32_Process -Filter "Name='wps.exe' OR Name='WINWORD.EXE'" |
            Where-Object { $before -notcontains $_.ProcessId -and $_.CommandLine -match '/Automation|-Embedding' } |
            ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    }
    if (-not (Test-Path $docx)) { throw "could not convert $TemplatePath to DOCX via $ComProgId" }
    "converted to DOCX via $ComProgId"
}
& "$s\probe_template.ps1" -TemplatePath $docx -OutDir $out | Out-Null
pwsh -NoProfile -STA -File "$s\export_pdf.ps1" -DocPath $docx -PdfPath (Join-Path $out 'template.pdf') -ComProgId $ComProgId | Out-Null
$py = Join-Path $env:LOCALAPPDATA 'md-to-journal-docx\venv\Scripts\python.exe'
if ((Test-Path $py) -and (Test-Path (Join-Path $out 'template.pdf'))) {
    & $py -I "$s\render_pages.py" (Join-Path $out 'template.pdf') (Join-Path $out 'pages') --pages all | Out-Null
}
"NO PROFILE for this template. Prepared in $out : template.docx, template_probe.md/.json, template.pdf, pages/"
"next: references/new_journal_profile.md (record original_doc_sha256 = $hash in template_visual_review.json)"
exit 2
