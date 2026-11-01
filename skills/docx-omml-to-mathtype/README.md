# docx-omml-to-mathtype

A Claude Code skill that converts the Word-native equations (OMML) in a `.docx` into real,
editable **MathType** objects (`Equation.DSMT4`) — **with WPS Office, no Microsoft Word license and
no VBA**.

把 `.docx` 中的 Word 原生公式（OMML）批量转换为可编辑的 MathType 公式对象。只需 WPS + MathType 7，
不需要 Word 授权，也不需要 VBA（WPS 个人版没有 VBA，MathType 的 “Convert Equations” 因此不可用）。

## Why another tool

| | this skill | MathType *Convert Equations* | Word-COM skills (e.g. word-omml-mathtype, thesis-mathtype-finalize) | omml-converter (PyPI) |
|---|---|---|---|---|
| Needs MS Word | no (WPS) | yes | yes | no |
| Needs VBA | no | yes | yes (MathType add-in) | no |
| Output | MathType 7 `Equation.DSMT4`, made by MathType itself | same | same | self-generated `Equation.DSMT6` (MathType 5) |
| Validation | per-equation render check, structural diff, PDF/PNG review | — | varies | — |

## Requirements

Windows 10/11 x64 · MathType 7 (tested 7.12.3) · WPS Office (tested 12.1.0.29162) · `OMML2MML.XSL`
from any Microsoft Office install (activation not required; not shipped here) · PowerShell 7 ·
uv or Python 3.10+. Full list and fixes: [`references/requirements.md`](references/requirements.md);
check with `pwsh -File scripts/check_runtime.ps1`.

## Install

Copy this folder to `~/.claude/skills/docx-omml-to-mathtype/` (Windows:
`%USERPROFILE%\.claude\skills\`). Claude Code picks it up automatically; ask it e.g.
"把 paper.docx 里的公式转成 MathType".

Without Claude:

```powershell
pwsh -File scripts\check_runtime.ps1
pwsh -File scripts\convert.ps1 -InputDocx paper.docx            # -> paper_mathtype.docx
```

## How it works

1. `prep_docx.py` replaces each OMML equation with a marker and converts it to MathML (Office's
   XSLT plus fixes for bold, script letters, spacing, `cases`, CJK text).
2. For each equation a worker process asks MathType's C API (`MT6.dll`) to turn the MathML on the
   clipboard into an OLE equation, renders it back to check it, and exits.
3. `wps_replace.ps1` pastes it over the marker through WPS COM and restarts MathType's API server.
4. Bookmarks are restored, the result is diffed against the original and rendered to PNG.

The many MathType/WPS quirks this works around are documented in
[`references/pitfalls.md`](references/pitfalls.md).

## Limitations

Only the main document body (not headers/footers/footnotes/text boxes). MS Word as host is
untested. Long display equations don't auto-break as MathType objects.

## Notices

MathType is a trademark of Wiris / Design Science; this project is not affiliated with them and
ships no MathType or Microsoft code. You need your own MathType and Office/WPS installations.
