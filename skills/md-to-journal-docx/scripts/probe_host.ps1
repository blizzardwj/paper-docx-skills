<#
Test, on a COPY of a real document, whether a COM host can do every step the pipeline needs:
open, walk story ranges, update fields, count pages/OMML, SaveAs DOCX and DOC, export PDF.
Registry presence is not proof (an unlicensed Word may still work; a registered host may not).

  pwsh -STA -File probe_host.ps1 -DocPath working.docx -ComProgId KWps.Application [-OutJson probe.json]
Exit code 0 = every capability passed.
#>
param([Parameter(Mandatory)] [string]$DocPath, [Parameter(Mandatory)] [string]$ComProgId, [string]$OutJson)
$tmp = Join-Path $env:TEMP ('md-to-journal-docx-hostprobe-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmp | Out-Null
$src = Join-Path $tmp 'probe_src.docx'; Copy-Item $DocPath $src
$before = @(Get-Process wps, winword -ErrorAction SilentlyContinue | ForEach-Object Id)
$cap = [ordered]@{ host = $ComProgId; version = $null; create = $false; foreign_documents = $null; open = $false; story_ranges = $false; fields_update = $false
                   page_count = $null; omath_count = $null; saveas_docx = $false; saveas_doc = $false; export_pdf = $false; errors = @() }
$app = $null
try {
    $app = New-Object -ComObject $ComProgId; $cap.create = $true
    $cap.version = "$($app.Name) $($app.Version) build $(try { $app.Build } catch { '?' })"
    $cap.foreign_documents = [int]$app.Documents.Count
    if ($cap.foreign_documents -ne 0) { throw 'instance already holds documents (in use by the user) - not touched' }
    $app.Visible = $false; try { $app.DisplayAlerts = 0 } catch {}
    $doc = $app.Documents.Open($src, $false, $false); $cap.open = $true
    try {
        $seen = [Collections.Generic.HashSet[string]]::new(); $r = $doc.StoryRanges.Item(1); $steps = 0
        while ($r -and $steps -lt 50) { $steps++; $key = "$($r.StoryType):$($r.Start):$($r.End)"; if (-not $seen.Add($key)) { break }; $r = $r.NextStoryRange }
        $cap.story_ranges = $true
    } catch { $cap.errors += "story ranges: $($_.Exception.Message)" }
    try { [void]$doc.Fields.Update(); $cap.fields_update = $true } catch { $cap.errors += "fields: $($_.Exception.Message)" }
    try { $cap.page_count = [int]$doc.ComputeStatistics(2) } catch { $cap.errors += "pages: $($_.Exception.Message)" }
    try { $cap.omath_count = [int]$doc.OMaths.Count } catch { $cap.errors += "OMaths: $($_.Exception.Message)" }
    foreach ($f in @(@('docx', 16), @('doc', 0))) {
        $p = Join-Path $tmp "probe_out.$($f[0])"
        try { $doc.SaveAs2($p, $f[1]) } catch { try { $doc.SaveAs($p, $f[1]) } catch { $cap.errors += "save $($f[0]): $($_.Exception.Message)" } }
        $cap["saveas_$($f[0])"] = Test-Path $p
    }
    $pdf = Join-Path $tmp 'probe_out.pdf'
    try { $doc.ExportAsFixedFormat($pdf, 17); $cap.export_pdf = Test-Path $pdf } catch { $cap.errors += "pdf: $($_.Exception.Message)" }
    $doc.Close(0)
} catch { $cap.errors += $_.Exception.Message }
finally {
    if ($app -and $cap.foreign_documents -eq 0) { try { $app.Quit() } catch {} }
    if ($app) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
    Start-Sleep -Seconds 2
    Get-CimInstance Win32_Process -Filter "Name='wps.exe' OR Name='WINWORD.EXE'" |
        Where-Object { $before -notcontains $_.ProcessId -and $_.CommandLine -match '/Automation|-Embedding' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
$cap.passed = $cap.create -and $cap.open -and $cap.story_ranges -and $cap.fields_update -and $cap.saveas_docx -and $cap.saveas_doc -and $cap.export_pdf -and ($cap.page_count -gt 0)
if ($OutJson) { $cap | ConvertTo-Json -Depth 4 | Set-Content $OutJson -Encoding utf8 }
[pscustomobject]$cap | Format-List | Out-String -Width 200
if (-not $cap.passed) { exit 1 }
