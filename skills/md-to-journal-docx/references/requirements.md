# Runtime requirements (dependencies)

`scripts/check_runtime.ps1` checks all of these and prints a fix for anything missing;
`make_docx.ps1` runs it first. "Tested" = versions the pipeline was verified with (2026-10-31).

## Required

| requirement | used by | tested | if missing |
|---|---|---|---|
| Windows 10/11 x64 | everything | Windows 11 10.0.26300 | not supported elsewhere |
| PowerShell 7+ (`pwsh`) | all `.ps1`; builder needs `#requires -Version 7`, `ConvertFrom-Json -AsHashtable`, `Process.Kill($true)` | 7.6.9 | `winget install Microsoft.PowerShell` |
| .NET C# compiler via `Add-Type` | `latex_to_omml.cs` (also used for readable formula text) | ships with PowerShell 7 | reinstall PowerShell 7 |
| Python env `%LOCALAPPDATA%\md-to-journal-docx\venv` with **pypandoc_binary** (bundles pandoc), **lxml**, **pymupdf** | pandoc formula engine (`tex2omml.py`), page renders (`render_pages.py`) | CPython 3.12, pandoc 3.9, lxml 6.1.3, PyMuPDF 2.0 | created automatically by `make_docx.ps1` with **uv** (`winget install astral-sh.uv`) or the `py` launcher; needs internet once |
| Document host with COM automation: **WPS Office** (`KWps.Application`, primary) or **desktop Word** (`Word.Application`) | `finalize_native.ps1` (field update, SaveAs DOCX/DOC, reopen, metrics), `export_pdf.ps1` | WPS Office 12.1.0.29162 personal edition; Word 16.0.20527 | install WPS Office for Windows; VBA is not needed |
| Interactive desktop session | host steps (COM) | console session | run from a logged-in desktop, not SSH or a service |
| Template mode only — journal profile: `template_reference.docx`, `format_mapping.json` (hash-bound to the template, `authoring_gate = passed`), `profile.json` | builder | one profile, validated (`validate_profile.ps1`) | `resolve_template.ps1` prepares the evidence; `references/new_journal_profile.md` |
| Fonts listed in the profile's `fonts_required` (plus Cambria Math for equations; plain mode: Cambria Math and a CJK font such as SimSun for Chinese text) | host rendering | Windows system and CJK fonts | install them; otherwise the host substitutes fonts and line/page breaks move |

## Host choice (WPS vs Word)

- Registration is not proof. Every `make_docx` run first runs `probe_host.ps1` on a copy of the
  built document (open, story ranges, field update, page/OMML count, SaveAs DOCX and DOC, PDF);
  with `-ComProgId auto` it tries WPS, then Word, and uses the first that passes. Results:
  `<work>/host_probe_<host>.json`.
- **WPS is tried first**: it passed every step (field update, SaveAs DOCX/DOC, reopen,
  page/table/equation counts, PDF).
- **Word works too**, including on a machine whose Microsoft 365 licence grace period has expired
  (COM still opened, edited, saved and exported). That is not guaranteed to last; don't rely on an
  unlicensed Word. Use `-ComProgId Word.Application` to force it.
- In testing both hosts produced identical page/section/table/equation counts. Line breaks can
  still differ between hosts and fonts; always review the pages.

## Not needed

- No Microsoft Office licence, no VBA, no MathType (MathType is the separate
  docx-omml-to-mathtype skill), no admin rights, no LaTeX installation.
- The template builder itself (`build_docx.ps1 -MathEngine cs`) needs neither Python nor a host;
  with the default pandoc engine it needs only the Python env. Plain mode always needs the Python
  env (pandoc); `-SkipHost` skips the host in both modes (no field update, save check or pages).
- A `.docx` template needs no host to be resolved; a `.doc`/`.dot` template without a profile is
  converted to DOCX through the host once.

## Created / touched by the skill

- Python env: `%LOCALAPPDATA%\md-to-journal-docx\venv`.
- Work files per run: `%TEMP%\md-to-journal-docx\<timestamp>\` (working DOCX, build report,
  content ledger, formula manifest, PDF, page PNGs) unless `-WorkDir` is given.
- Outputs only in `-OutputDirectory` (`<stem>.docx`, `<stem>.doc`, `<stem>_save_check.json`);
  existing files are replaced only with `-Force`. Inputs and the profile are never written.
- A hidden WPS/Word automation instance is started and closed; documents the user has open are
  left alone (the run refuses to use a COM instance that already holds documents).
