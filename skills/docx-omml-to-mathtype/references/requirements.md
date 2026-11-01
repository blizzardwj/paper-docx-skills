# Runtime requirements

`scripts/check_runtime.ps1` checks all of these and prints a fix for anything missing.
"Tested" = the versions the pipeline was verified with.

## Required

| requirement | why | where / tested | if missing |
|---|---|---|---|
| Windows 10/11 x64 | MathType API, COM automation, Windows clipboard | Windows 11 (10.0.26300) | not supported elsewhere |
| Interactive desktop session | the run uses the clipboard and starts WPS/MathType | console session | run from a logged-in desktop, not SSH or a service |
| PowerShell 7+ (`pwsh`) | scripts use `Start-Process -PassThru`, `Process.Kill($true)`, `-STA` | 7.6.9 | `winget install Microsoft.PowerShell` |
| MathType 7 desktop (full install, licensed or trial) | does the actual conversion | 7.12.3 | install from wiris.com; the fonts-only package is not enough |
| `MT6.dll` 64-bit | MathType C API: `MTAPIConnect`, `MTXFormEqn`, `MTXFormSetPrefs`, `MTGetPrefsFromFile`, `MTEquationOnClipboard` | `C:\Program Files (x86)\MathType\System\64\MT6.dll` | reinstall MathType to the default folder |
| `MathType.exe` (`-mtrpc` API server) | started automatically by `MTAPIConnect` | `C:\Program Files (x86)\MathType\MathType.exe` | part of MathType |
| `Times+Symbol 12.eqp` | base equation preferences (font size/text font are patched per run) | `C:\Program Files (x86)\MathType\Preferences\` | part of MathType |
| `Equation.DSMT4` OLE class registered | WPS/Word embed and later re-open the objects | `HKCR\Equation.DSMT4` | repair MathType |
| WPS Office (COM `KWps.Application`) — or desktop Word (`Word.Application`, untested) | opens the document, finds markers, pastes the objects, exports PDF | WPS Office 12.1.0.29162 personal edition | install WPS for Windows; VBA is **not** needed |
| `OMML2MML.XSL` | Microsoft's OMML→MathML stylesheet used by `prep_docx.py` | `C:\Program Files\Microsoft Office\root\Office16\` (Microsoft 365) | ships only with Microsoft Office — any install works, activation not required; WPS does not include it. Install Office/trial, or pass `-Xsl <path>` to a copy |
| uv (or Python 3.10+ with the `py` launcher) | builds the cached Python env once | uv 0.12.17, CPython 3.12 | `winget install astral-sh.uv` |
| Internet on first run | downloads `lxml` and `pymupdf` into the env | — | afterwards fully offline |
| Fonts: Times New Roman, Symbol, MT Extra, plus the CJK text font (default SimSun/宋体) | MathType equation styles; CJK inside `\text{}` | Windows + MathType fonts | install, or choose another CJK font with `-CjkFont` |

## Created by the skill

- Python env: `%LOCALAPPDATA%\docx-omml-to-mathtype\venv` (lxml, pymupdf).
- Work files per run: `%TEMP%\docx-omml-to-mathtype\<timestamp>\` (marked docx, `eqs.json`,
  per-equation worker logs and GIFs, `prefs.eqp`, PDF, page PNGs).

## Side effects during a run

- The clipboard is overwritten many times — don't copy/paste meanwhile.
- MathType's API server (`MathType.exe -mtrpc`) and its own `MathTypeLib.exe` child are killed
  and restarted after every equation. A MathType editor the user opened is a different process
  and is left alone, but saving MathType work before the run is still wise.
- A hidden WPS automation instance is started and closed; WPS windows the user has open are left
  alone, and only `/Automation` processes started by the run are cleaned up.

## Not supported / untested

- Equations in headers, footers, footnotes, endnotes, text boxes (only `word/document.xml`; a
  warning is printed when other parts contain equations).
- MS Word as host (`-ComProgId Word.Application`) — same object model, not tested. With licensed
  Word + MathType, MathType's own *Convert Equations* is the simpler route.
- MathType 6 / MathType Lite; macOS.
