# Puts ONE validated MathType equation on the clipboard, then exits.
# Runs in its own process because MT6.dll occasionally crashes the host (0xC0000005).
# Last stdout line: RESULT ok=<True|False> w=<px> h=<px> tries=<n>
param(
    [Parameter(Mandatory)] [string]$EqJson,
    [Parameter(Mandatory)] [int]$Eq,
    [Parameter(Mandatory)] [string]$PrefsFile,
    [Parameter(Mandatory)] [string]$GifDir,
    [int]$MaxTries = 6
)
Add-Type -Path "$PSScriptRoot\MtApi.cs"
Add-Type -AssemblyName System.Drawing, System.Windows.Forms

function Set-MathMLClipboard([string]$mml) {
    # MathType recognises MathML more reliably when the registered MathML formats are present too
    $dob = New-Object System.Windows.Forms.DataObject
    foreach ($fmt in "MathML", "MathML Presentation", "application/mathml+xml") {
        $dob.SetData($fmt, $false, (New-Object IO.MemoryStream (, [Text.Encoding]::UTF8.GetBytes($mml))))
    }
    $dob.SetText($mml)
    [System.Windows.Forms.Clipboard]::SetDataObject($dob, $true, 5, 100)
}

$e = (Get-Content $EqJson -Raw -Encoding utf8 | ConvertFrom-Json) | Where-Object index -eq $Eq
$plainLen = ($e.mathml -replace '<[^>]+>', '').Length
$gif = Join-Path $GifDir ("eq{0:D3}.gif" -f $Eq)
[void][MtApi]::MTAPIConnect(0, 30)
$prefs = [MtApi]::LoadPrefs($PrefsFile)
$ok = $false; $w = 0; $h = 0
for ($try = 1; $try -le $MaxTries -and -not $ok; $try++) {
    Set-MathMLClipboard $e.mathml
    Start-Sleep -Milliseconds (100 * $try)
    $rc = [MtApi]::ConvertClipboardMathML($prefs, [MtApi]::CLIPBOARD, [MtApi]::PICT, $null)
    $onClip = [MtApi]::MTEquationOnClipboard()
    if ($rc -ne 0 -or $onClip -ne 8) { "try $try convert rc=$rc onClip=$onClip"; continue }
    # Verify the clipboard copy itself. MathType sometimes returns 0 but typesets the MathML
    # source as literal text (~6 px per char); a real equation is far narrower.
    Remove-Item $gif -ErrorAction SilentlyContinue
    $vrc = [MtApi]::RenderClipboardEquation($prefs, [MtApi]::GIF, $gif)
    if ($vrc -ne 0 -or -not (Test-Path $gif)) { "try $try verify rc=$vrc"; continue }
    $img = [System.Drawing.Image]::FromFile($gif); $w = $img.Width; $h = $img.Height; $img.Dispose()
    if ($w -gt [Math]::Max(60, 25 * $plainLen)) { "try $try literal-text render (w=$w)"; continue }
    $ok = $true
}
[void][MtApi]::MTAPIDisconnect()
"RESULT ok=$ok w=$w h=$h tries=$($try - 1)"
