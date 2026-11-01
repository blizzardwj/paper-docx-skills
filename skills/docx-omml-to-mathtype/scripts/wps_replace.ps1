# Replace QQMTEQnnnQQ markers in a .docx with MathType equations, driving WPS (or Word) over COM.
# Each equation: worker process -> validated OLE equation on clipboard -> Selection.Paste -> kill MathType server.
param(
    [Parameter(Mandatory)] [string]$DocPath,
    [Parameter(Mandatory)] [string]$EqJson,
    [Parameter(Mandatory)] [string]$PrefsFile,
    [Parameter(Mandatory)] [string]$WorkDir,
    [string]$ComProgId = "KWps.Application",
    [int]$MaxProcs = 6,
    [int]$WorkerTimeoutSec = 90
)
$ErrorActionPreference = "Stop"
$s = $PSScriptRoot
$gifDir = Join-Path $WorkDir "gif"
New-Item -ItemType Directory -Force $gifDir | Out-Null
$eqs = Get-Content $EqJson -Raw -Encoding utf8 | ConvertFrom-Json | Where-Object converted
"to convert: $($eqs.Count)"

function Stop-MathTypeServer {
    # The API server (MathType.exe -mtrpc) accumulates hidden equation windows until it shows a
    # modal "too many windows" dialog that blocks everything; restart it for every equation.
    # Only the API server and its own helper; a MathType editor the user opened is left alone.
    $servers = @(Get-CimInstance Win32_Process -Filter "Name='MathType.exe'" | Where-Object { $_.CommandLine -match '-mtrpc' } | ForEach-Object ProcessId)
    if (-not $servers) { return }
    Get-CimInstance Win32_Process -Filter "Name='MathTypeLib.exe'" | Where-Object { $servers -contains $_.ParentProcessId } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    $servers | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }
}

function Invoke-Worker($index) {
    $log = Join-Path $gifDir ("worker_eq{0:D3}.log" -f $index)
    $p = Start-Process pwsh -PassThru -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError "$log.err" `
        -ArgumentList @("-NoProfile", "-STA", "-File", "`"$s\mt_worker.ps1`"", "-EqJson", "`"$EqJson`"",
                        "-Eq", $index, "-PrefsFile", "`"$PrefsFile`"", "-GifDir", "`"$gifDir`"")
    if (-not $p.WaitForExit($WorkerTimeoutSec * 1000)) { $p.Kill($true); return "TIMEOUT" }
    return "$(Get-Content $log -ErrorAction SilentlyContinue | Select-Object -Last 1)"
}

$wpsBefore = @(Get-Process wps, winword -ErrorAction SilentlyContinue | ForEach-Object Id)
$app = New-Object -ComObject $ComProgId
$app.Visible = $false
try { $app.DisplayAlerts = 0 } catch {}
$foreignDocs = $app.Documents.Count      # >0 means COM attached to an instance the user is using
$failed = 0
try {
    Stop-MathTypeServer
    $doc = $app.Documents.Open($DocPath)
    foreach ($e in $eqs) {
        $kind = if ($e.display) { "display" } else { "inline" }
        $res = $null
        for ($p = 1; $p -le $MaxProcs; $p++) {
            $last = Invoke-Worker $e.index
            if ($last -match 'RESULT ok=True w=(\d+) h=(\d+)') { $res = @{ w = [int]$Matches[1]; h = [int]$Matches[2]; procs = $p }; break }
            Stop-MathTypeServer
        }
        if (-not $res) { $failed++; "eq {0,3} ({1}): FAILED after {2} worker processes (marker left in place)" -f $e.index, $kind, $MaxProcs; continue }

        $rng = $doc.Content
        if (-not $rng.Find.Execute($e.marker)) { $failed++; "eq {0,3}: marker not found" -f $e.index; Stop-MathTypeServer; continue }
        $before = $doc.InlineShapes.Count
        $rng.Select()                 # WPS: Range.Paste() throws; Selection.Paste() works
        $app.Selection.Paste()
        Stop-MathTypeServer
        $markerLeft = $doc.Content.Text.Contains($e.marker)
        if ($markerLeft -or $doc.InlineShapes.Count -ne $before + 1) { $failed++ }
        "eq {0,3} ({1}): workers={2} gif={3}x{4} shapes {5}->{6} marker_left={7}" -f `
            $e.index, $kind, $res.procs, $res.w, $res.h, $before, $doc.InlineShapes.Count, $markerLeft
    }
    $dsmt = 0; foreach ($sh in $doc.InlineShapes) { try { if ($sh.OLEFormat.ProgID -like "Equation.DSMT*") { $dsmt++ } } catch {} }
    "MathType objects in document: $dsmt; problems: $failed"
    $doc.Save()
    $doc.Close()
    "saved"
}
finally {
    Stop-MathTypeServer
    if ($foreignDocs -eq 0) { try { $app.Quit() } catch {} }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    Start-Sleep -Seconds 2
    # clean up only automation instances this run started
    Get-CimInstance Win32_Process -Filter "Name='wps.exe' OR Name='WINWORD.EXE'" |
        Where-Object { $wpsBefore -notcontains $_.ProcessId -and $_.CommandLine -match '/Automation|-Embedding' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
}
if ($failed) { exit 2 }
