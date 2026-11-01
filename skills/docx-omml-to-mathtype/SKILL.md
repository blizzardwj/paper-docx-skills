---
name: docx-omml-to-mathtype
description: Convert the Word-native equations (OMML, the ones made by Word's/WPS's built-in equation editor or by pandoc/Markdown export) in a .docx into real MathType equation objects (Equation.DSMT4), on Windows with MathType 7 installed — works with WPS only, no Word license or VBA needed. Use this whenever the user wants formulas in a Word/WPS document turned into MathType format, e.g. because a journal requires MathType, MathType's "Convert Equations" says "install VBA first", or they ask "把文档里的公式转成 MathType", "公式是不是 mathtype 格式", "批量转换公式", even if they don't name OMML.
compatibility: Windows 10/11 x64, MathType 7, WPS Office (or desktop Word), Microsoft Office's OMML2MML.XSL, PowerShell 7, uv or Python 3.10+. Details in references/requirements.md.
---

# DOCX OMML → MathType

Turns every Word-native equation in a .docx into an editable MathType OLE object while keeping
body text, bookmarks and inline/display placement. The pipeline is fully scripted and already
debugged — run the scripts; don't re-derive the approach.

## Where things are

| need | read |
|---|---|
| what must be installed, how to check, how to fix | `references/requirements.md` |
| a step failed / output looks wrong / about to change a script | `references/pitfalls.md` |
| what changed between versions | `CHANGELOG.md` |

Keep this file an outline. New problems and their fixes go into `references/pitfalls.md`, a line
in `CHANGELOG.md`, and the scripts — not here, unless the workflow itself changes.

## Workflow

1. **Check the runtime** — `pwsh -File scripts/check_runtime.ps1`. If anything is MISSING, tell the
   user what to install (the script prints the fix) and stop.
2. **Inspect the document** — count OMML equations (inline and display). None → say so and stop.
3. **Protect the user's file** — output goes to a new file (`<input>_mathtype.docx` by default);
   the input is never written. Ask the user not to use the clipboard during the run.
4. **Convert** — `pwsh -File scripts/convert.ps1 -InputDocx <file> [-OutputDocx <out>]`.
   About 5–10 s per equation; run it in the background for long documents.
   Use `-Only 1,7` to pilot a few equations on an unfamiliar document first.
5. **Verify** — every line of the verification block must be `[OK]` (exit code 0). On failure,
   look the symptom up in `references/pitfalls.md`; random MathType failures usually pass on rerun.
6. **Review visually** — look at the page renders in the work dir (`pages/pageNN.png`): display
   equations, bold, script letters, braces, CJK text, inline baseline alignment.
7. **Report** — equation counts (inline / display), verification result, output path; ask the user
   to double-click one equation in WPS to confirm MathType opens it. Offer, don't apply, fixes for
   layout problems inherited from the original (e.g. a display equation wider than its column).

## Scripts (`scripts/`)

| script | role |
|---|---|
| `check_runtime.ps1` | preflight of every requirement, with install hints |
| `convert.ps1` | one-shot pipeline: preflight → prep → replace → bookmarks → verify → PDF/PNG |
| `prep_docx.py` | OMML → text markers + MathType-safe MathML (`eqs.json`) |
| `wps_replace.ps1` | COM driver: one worker per equation, paste over marker, restart MathType server |
| `mt_worker.ps1`, `MtApi.cs` | one equation → validated MathType OLE object on the clipboard |
| `restore_bookmarks.py` | re-adds bookmarks removed by pasting |
| `verify_docx.py` | structural checks (exit 0 = pass) |
| `export_pdf.ps1`, `render_pages.py` | PDF export and PNG renders for review |
