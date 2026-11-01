# Export a .docx to PDF through WPS (or Word) COM, read-only, for visual review.
param([Parameter(Mandatory)] [string]$DocPath, [Parameter(Mandatory)] [string]$PdfPath, [string]$ComProgId = "KWps.Application")
$before = @(Get-Process wps, winword -ErrorAction SilentlyContinue | ForEach-Object Id)
$app = New-Object -ComObject $ComProgId
$app.Visible = $false
try { $app.DisplayAlerts = 0 } catch {}
$foreign = $app.Documents.Count
try {
    $doc = $app.Documents.Open($DocPath, $false, $true)
    $doc.ExportAsFixedFormat($PdfPath, 17)
    $doc.Close(0)
    "exported: $(Test-Path $PdfPath)"
} finally {
    if ($foreign -eq 0) { try { $app.Quit() } catch {} }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    Start-Sleep -Seconds 2
    Get-CimInstance Win32_Process -Filter "Name='wps.exe' OR Name='WINWORD.EXE'" |
        Where-Object { $before -notcontains $_.ProcessId -and $_.CommandLine -match '/Automation|-Embedding' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
}
