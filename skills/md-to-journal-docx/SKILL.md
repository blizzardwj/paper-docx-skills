---
name: md-to-journal-docx
description: Convert a Markdown paper/thesis/report into an editable DOCX with Word-native (OMML) equations — either plainly (no template: bundled neutral styles, Times New Roman/宋体, black bold headings) or typeset into a journal's Word template (template-cloned styles and sections, single/multi-column layout, three-line tables, captions, references, page-number fields) — then native-save it through WPS or Word, reopen, count pages/tables/equations and render every page for review. Windows. Use this whenever the user wants Markdown turned into .docx/.doc, with or without a template, rebuilt after editing the Markdown, or moved to another journal's template, e.g. "把 Markdown 转成 Word", "把 Markdown 稿排进期刊模板", "从 md 生成投稿 docx", "按这个 Word 模板排版", even if they don't name the template. MathType conversion is a separate step (docx-omml-to-mathtype skill).
compatibility: Windows 10/11 x64, PowerShell 7, WPS Office or desktop Word (COM), uv or Python 3.10+ (pandoc via pypandoc_binary, lxml, pymupdf). Details in references/requirements.md.
---

# Markdown → DOCX (native equations), with or without a template

Two inputs decide everything:

| inputs | mode | result |
|---|---|---|
| Markdown only | **plain** | bundled style reference (`assets/plain_reference[_zh].docx`); equations, tables, pictures, headings converted as written |
| Markdown + template (`.docx/.doc/.dotx`) | **template** | the template's styles, sections, headers/footers, journal rules, via the template's **profile** |

A profile is what the skill learns from a template once (role → sample paragraph, sizes, sections,
how manuscripts for it are written). It is found by the template file's SHA-256; a template without
a profile gets one made first (probe → decide → validate → user approves), then it is reused.
Everything specific to a journal or paper is data (profile, optional per-paper config); each
uncertain fact is settled by a script before it is relied on.

## Where things are

| need | read |
|---|---|
| install / check the runtime | `references/requirements.md` |
| how Markdown is recognised (template mode), all rule keys, per-paper config files | `references/rules.md`, `references/inputs.md` |
| a template has no profile yet | `references/new_journal_profile.md` |
| a step failed / output looks wrong / before changing a script | `references/pitfalls.md` |
| what changed | `CHANGELOG.md` |
| existing profiles and their conventions | `profiles/<name>/profile.md` |

Keep this file an outline. New problems → `references/pitfalls.md` + `CHANGELOG.md` (+ scripts);
journal- or paper-specific conventions → the profile or the paper config, never the scripts or
this file.

## Uncertainties and how each becomes certain

| uncertain | settled by | mode |
|---|---|---|
| runtime (shell, Python env, fonts, COM registration) | `check_runtime.ps1` | both, every run |
| can this host really open/update/save/export? | `probe_host.ps1` on a copy of the built document; auto mode falls back to the next host | both, every run |
| does the Markdown convert completely? | plain: `plain_docx.py` fails on unconverted LaTeX, `$` left as text, raw LaTeX that would be dropped, missing pictures, equation count mismatch; template: `build_docx.ps1 -ValidateOnly` (every line mapped to a block, every formula converted, all failures listed at once) | both, every run |
| which profile belongs to this template? | `resolve_template.ps1` (SHA-256 of the file vs every profile) | template |
| what the template contains | `resolve_template.ps1` → `probe_template.ps1` facts + rendered pages → agent decisions | template, no profile yet |
| are the profile's decisions true for the template? | `validate_profile.ps1` | template, after writing/changing a profile or paper overrides |
| did content survive? | content ledger + preserved-part hashes (template build); counts after native save | both |
| does it look right? | page PNGs reviewed by the agent; fixed through config, then rebuilt | both |

## Workflow

1. **Runtime** — `pwsh -File scripts/check_runtime.ps1`. MISSING → tell the user what to install
   (the script prints the fix) and stop.
2. **Mode** — no template from the user → plain; go to 4. Template given → step 3.
3. **Profile** (template mode) — `make_docx` resolves it from `-TemplatePath`. Exit 2 = no profile:
   the evidence is prepared in `<work>/template/` → `references/new_journal_profile.md` (needs the
   user's approval of the template pages) → `validate_profile.ps1` 0 FAIL → run again. Then
   compare the manuscript with the profile's rules (`profile.md`); where this paper writes things
   differently, put rule overrides in its `paper_overrides.json` (`references/rules.md`) instead of
   editing the manuscript or the profile, and re-run `validate_profile.ps1 -PaperOverridesPath …`.
4. **Build + host check + save + render** —
   `pwsh -File scripts/make_docx.ps1 -ManuscriptPath <md> [-TemplatePath <template>] [-PaperConfigDir <cfg>] -OutputDirectory <dir> -OutputStem <stem>`.
   Use a new stem; never overwrite the user's deliverables. Plain-mode errors name the offending
   text: fix the Markdown (or `-AllowLiteralDollar` when the text really has `$` signs).
5. **Read the summary** — host probe passed, counts match the source, no table spans pages,
   exit 0. Failures: `references/pitfalls.md`.
6. **Review pages** (`<work>/pages/pageNN.png`): headings, every display equation, tables,
   pictures, references, last page; in template mode also title block, headers/footers,
   figure–caption adjacency. Template mode: fix layout through `layout_config.json` / overrides and
   rebuild, not by hand-editing the DOCX.
7. **Report** — mode, output paths (DOCX master, DOC copy, save-check JSON), counts, host used,
   what was reviewed, remaining placeholders and open WARNs. DOC turns equations into pictures; the
   DOCX is the master. For MathType, run the docx-omml-to-mathtype skill on the DOCX next.

## Scripts (`scripts/`)

| script | role |
|---|---|
| `make_docx.ps1` | one-shot, both modes: check → (resolve profile) → build → host probe → native save → PDF/PNG |
| `check_runtime.ps1` | platform, Python env, COM registration, profile integrity, fonts |
| `plain_docx.py` | plain mode: bundled styles, equation font fixes, table/picture layout, completeness checks |
| `make_plain_reference.py` | rebuilds `assets/plain_reference[_zh].docx` from pandoc's default (after a pandoc upgrade or a style change) |
| `resolve_template.ps1` | template → profile by SHA-256; none → DOCX copy, probe, rendered pages, exit 2 |
| `probe_template.ps1` | facts of any template → `template_probe.json` / `.md` |
| `validate_profile.ps1` | profile (+ paper overrides) checked against its template |
| `probe_host.ps1` | proves a COM host can do every host step on a copy of the document |
| `build_docx.ps1` | template mode: Markdown + profile (+ paper config) → DOCX, ledger, formula manifest (no COM) |
| `default_rules.json` | neutral defaults for every rule (profiles/papers override) |
| `tex2omml.py` / `latex_to_omml.cs` | formula engines: pandoc/texmath (default) / legacy strict subset |
| `finalize_native.ps1` | host save: fields, SaveAs DOCX + DOC, reopen read-only, metrics JSON |
| `export_pdf.ps1`, `render_pages.py` | PDF through the host, PNG per page |
