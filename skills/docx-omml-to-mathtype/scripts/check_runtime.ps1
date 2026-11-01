<#
Preflight: report every runtime requirement of this skill as OK / MISSING / WARN, with how to fix.
Exit code 0 = all required items present.

  pwsh -File check_runtime.ps1 [-ComProgId KWps.Application] [-CjkFont SimSun] [-Xsl path]
#>
param([string]$ComProgId = "KWps.Application", [string]$CjkFont = "SimSun", [string]$Xsl)
$results = @()
function Add-Check($name, $status, $detail, $fix = "") { $script:results += [pscustomobject]@{ Status = $status; Item = $name; Detail = $detail; Fix = $fix } }
function Get-Uninstall($pattern) {
    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*, HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*, HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match $pattern } | Select-Object -First 1
}

# --- OS / shell / session ---
$os = [Environment]::OSVersion
Add-Check "Windows x64" $(if ($IsWindows -and [Environment]::Is64BitOperatingSystem) { "OK" } else { "MISSING" }) "$($os.VersionString)" "Windows 10/11 64-bit only (MathType API, COM, clipboard)"
Add-Check "PowerShell 7+" $(if ($PSVersionTable.PSVersion.Major -ge 7) { "OK" } else { "MISSING" }) "$($PSVersionTable.PSVersion)" "winget install Microsoft.PowerShell"
$interactive = [Environment]::UserInteractive -and $env:SESSIONNAME -ne "Services"
Add-Check "interactive desktop session" $(if ($interactive) { "OK" } else { "MISSING" }) "session=$env:SESSIONNAME" "run from a logged-in desktop (clipboard + COM); not over SSH or as a service"

# --- MathType ---
$mt = Get-Uninstall '^MathType'
$mtDll = "C:\Program Files (x86)\MathType\System\64\MT6.dll"
$eqp = "C:\Program Files (x86)\MathType\Preferences\Times+Symbol 12.eqp"
Add-Check "MathType 7 installed" $(if ($mt -and $mt.DisplayVersion -like "7.*") { "OK" } elseif ($mt) { "WARN" } else { "MISSING" }) "$($mt.DisplayName) $($mt.DisplayVersion) (tested: 7.12.3)" "install MathType 7 desktop (licensed or trial; full install, not fonts-only)"
Add-Check "MathType 64-bit API (MT6.dll)" $(if (Test-Path $mtDll) { "OK" } else { "MISSING" }) $mtDll "reinstall MathType 7 to the default folder"
Add-Check "MathType preferences file" $(if (Test-Path $eqp) { "OK" } else { "MISSING" }) $eqp "reinstall MathType 7 (Preferences folder)"
Add-Check "Equation.DSMT4 OLE class" $(if (Test-Path "Registry::HKEY_CLASSES_ROOT\Equation.DSMT4") { "OK" } else { "MISSING" }) "registry HKCR\Equation.DSMT4" "repair/reinstall MathType 7"

# --- document host (WPS or Word) ---
$host_ = if ($ComProgId -like "KWps*") { Get-Uninstall '^WPS Office' } else { Get-Uninstall 'Microsoft (Office|365)' }
Add-Check "COM host $ComProgId" $(if (Test-Path "Registry::HKEY_CLASSES_ROOT\$ComProgId") { "OK" } else { "MISSING" }) "$($host_.DisplayName) $($host_.DisplayVersion) (tested: WPS Office 12.1.0.29162)" $(if ($ComProgId -like "KWps*") { "install WPS Office for Windows (personal edition is enough; VBA not needed)" } else { "install desktop Microsoft Word (untested path)" })

# --- OMML -> MathML stylesheet (ships only with Microsoft Office) ---
if (-not $Xsl) {
    $Xsl = @("C:\Program Files\Microsoft Office\root\Office16\OMML2MML.XSL", "C:\Program Files (x86)\Microsoft Office\root\Office16\OMML2MML.XSL",
             "C:\Program Files\Microsoft Office\Office16\OMML2MML.XSL", "C:\Program Files (x86)\Microsoft Office\Office16\OMML2MML.XSL") |
        Where-Object { Test-Path $_ } | Select-Object -First 1
}
Add-Check "OMML2MML.XSL" $(if ($Xsl -and (Test-Path $Xsl)) { "OK" } else { "MISSING" }) "$Xsl" "comes with Microsoft Office (any install, activation not needed); WPS does not ship it. Install Office (or the free trial), or pass -Xsl <path> to a copy of the file from a machine with Office"

# --- Python env for lxml/pymupdf ---
$venvPy = Join-Path $env:LOCALAPPDATA "docx-omml-to-mathtype\venv\Scripts\python.exe"
$haveEnv = $false
if (Test-Path $venvPy) { & $venvPy -I -c "import lxml, pymupdf" 2>$null; $haveEnv = ($LASTEXITCODE -eq 0) }
$uv = Get-Command uv -ErrorAction SilentlyContinue; $pyl = Get-Command py -ErrorAction SilentlyContinue
if ($haveEnv) { Add-Check "Python env (lxml, pymupdf)" "OK" $venvPy }
elseif ($uv -or $pyl) { Add-Check "Python env (lxml, pymupdf)" "WARN" "not created yet; first run builds it with $(if ($uv) { 'uv' } else { 'py launcher' }) (needs internet)" "" }
else { Add-Check "Python env (lxml, pymupdf)" "MISSING" "neither uv nor the py launcher found" "winget install astral-sh.uv   (or install Python 3.10+ with the py launcher)" }

# --- fonts ---
Add-Type -AssemblyName System.Drawing
foreach ($f in @("Times New Roman", "Symbol", "MT Extra", $CjkFont)) {
    $ok = try { [void](New-Object System.Drawing.FontFamily $f); $true } catch { $false }   # resolves localized names (SimSun = 宋体)
    Add-Check "font $f" $(if ($ok) { "OK" } else { "MISSING" }) "" $(if ($f -eq "MT Extra") { "installed by MathType" } elseif ($f -eq $CjkFont) { "install the font or pass -CjkFont <installed CJK font>" } else { "Windows system font" })
}

$results | Format-Table Status, Item, Detail -AutoSize -Wrap | Out-String -Width 200
$missing = $results | Where-Object Status -eq "MISSING"
if ($missing) {
    "How to fix:"
    $missing | ForEach-Object { "  - $($_.Item): $($_.Fix)" }
    exit 1
}
"runtime OK"
