<#
Preflight for md-to-journal-docx: report each requirement as OK / WARN / MISSING with a fix.
Exit code 0 = everything required for the requested steps is present.

  pwsh -File check_runtime.ps1 [-ProfileDir <profile>] [-MathEngine pandoc|cs] [-ComProgId KWps.Application|Word.Application|auto] [-SkipHost]
#>
param([string]$ProfileDir, [ValidateSet('pandoc','cs')] [string]$MathEngine = 'pandoc', [string]$ComProgId = 'auto', [switch]$SkipHost)
$results = @()
function Add-Check($name, $status, $detail, $fix = "") { $script:results += [pscustomobject]@{ Status = $status; Item = $name; Detail = $detail; Fix = $fix } }
function Get-Uninstall($pattern) {
    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*, HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match $pattern } | Select-Object -First 1
}

# --- platform ---
Add-Check "Windows x64" $(if ($IsWindows -and [Environment]::Is64BitOperatingSystem) { "OK" } else { "MISSING" }) ([Environment]::OSVersion.VersionString) "Windows 10/11 64-bit (builder uses .NET; host steps use COM)"
Add-Check "PowerShell 7+" $(if ($PSVersionTable.PSVersion.Major -ge 7) { "OK" } else { "MISSING" }) "$($PSVersionTable.PSVersion)" "winget install Microsoft.PowerShell"
Add-Check ".NET C# compiler (Add-Type)" $(try { Add-Type -TypeDefinition 'public static class MtjProbe { public static int One() { return 1; } }' -ErrorAction Stop; "OK" } catch { if ($_.Exception.Message -match 'already exists') { "OK" } else { "MISSING" } }) "compiles latex_to_omml.cs" "comes with PowerShell 7"

# --- python env (pandoc engine, page renders) ---
$venvPy = Join-Path $env:LOCALAPPDATA "md-to-journal-docx\venv\Scripts\python.exe"
$envDetail = "not created"
$envOk = $false
if (Test-Path $venvPy) {
    $probe = & $venvPy -I -c "import pypandoc, lxml, pymupdf; print(pypandoc.get_pandoc_version())" 2>$null
    $envOk = ($LASTEXITCODE -eq 0); $envDetail = if ($envOk) { "$venvPy (pandoc $probe)" } else { "$venvPy (packages missing)" }
}
$uv = Get-Command uv -ErrorAction SilentlyContinue; $pyl = Get-Command py -ErrorAction SilentlyContinue
$needPy = ($MathEngine -eq 'pandoc') -or -not $SkipHost
if ($envOk) { Add-Check "Python env (pypandoc_binary, lxml, pymupdf)" "OK" $envDetail }
elseif ($uv -or $pyl) { Add-Check "Python env (pypandoc_binary, lxml, pymupdf)" $(if ($needPy) { "WARN" } else { "OK" }) "$envDetail; make_docx.ps1 creates it with $(if ($uv) { 'uv' } else { 'py' }) (needs internet once)" }
else { Add-Check "Python env (pypandoc_binary, lxml, pymupdf)" $(if ($needPy) { "MISSING" } else { "WARN" }) "neither uv nor py launcher" "winget install astral-sh.uv   (or Python 3.10+ with the py launcher)" }

# --- document host for native save / field update / PDF ---
if (-not $SkipHost) {
    $wps = Test-Path "Registry::HKEY_CLASSES_ROOT\KWps.Application"
    $word = Test-Path "Registry::HKEY_CLASSES_ROOT\Word.Application"
    $wpsApp = Get-Uninstall '^WPS Office'; $officeApp = Get-Uninstall 'Microsoft (Office|365)'
    $chosen = if ($ComProgId -ne 'auto') { $ComProgId } elseif ($wps) { 'KWps.Application' } elseif ($word) { 'Word.Application' } else { $null }
    $chosenOk = $chosen -and (Test-Path "Registry::HKEY_CLASSES_ROOT\$chosen")
    Add-Check "document host (COM)" $(if ($chosenOk) { "OK" } else { "MISSING" }) $(if ($chosen) { "$chosen  | WPS: $($wpsApp.DisplayVersion) | Office: $($officeApp.DisplayVersion) (tested: WPS 12.1.0.29162; Word 16.0.20527)" } else { "neither KWps.Application nor Word.Application registered" }) "install WPS Office for Windows (personal edition is enough) - or desktop Word"
    $interactive = [Environment]::UserInteractive -and $env:SESSIONNAME -ne "Services"
    Add-Check "interactive desktop session" $(if ($interactive) { "OK" } else { "MISSING" }) "session=$env:SESSIONNAME" "run host steps from a logged-in desktop, not SSH/service"
    "HOST=$chosen" | Out-Null
}

# --- journal profile ---
if ($ProfileDir) {
    $tpl = Join-Path $ProfileDir 'template_reference.docx'; $map = Join-Path $ProfileDir 'format_mapping.json'; $prof = Join-Path $ProfileDir 'profile.json'
    $have = (Test-Path $tpl) -and (Test-Path $map) -and (Test-Path $prof)
    $detail = $ProfileDir; $status = if ($have) { "OK" } else { "MISSING" }
    if ($have) {
        $m = Get-Content $map -Raw | ConvertFrom-Json -Depth 30
        if ($m.reference_sha256 -ne (Get-FileHash $tpl).Hash) { $status = "MISSING"; $detail += " (format_mapping.json does not match template_reference.docx)" }
        elseif ($m.authoring_gate -ne 'passed') { $status = "MISSING"; $detail += " (authoring_gate != passed: template not visually approved)" }
    }
    Add-Check "journal profile" $status $detail "see references/new_journal_profile.md"
    if ($have) {
        Add-Type -AssemblyName System.Drawing
        foreach ($f in @((Get-Content $prof -Raw -Encoding utf8 | ConvertFrom-Json).fonts_required)) {
            $ok = try { [void](New-Object System.Drawing.FontFamily $f); $true } catch { $false }   # resolves localized names (SimSun = 宋体)
            Add-Check "font $f" $(if ($ok) { "OK" } else { "WARN" }) "" "install the font; the host substitutes a different one and line breaks change"
        }
    }
}

$results | Format-Table Status, Item, Detail -AutoSize -Wrap | Out-String -Width 220
$missing = $results | Where-Object Status -eq "MISSING"
if ($missing) { "How to fix:"; $missing | ForEach-Object { "  - $($_.Item): $($_.Fix)" }; exit 1 }
"runtime OK"
