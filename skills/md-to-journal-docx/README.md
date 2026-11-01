# md-to-journal-docx

A Claude Code skill that turns a Markdown manuscript into an editable DOCX with Word-native
(OMML) equations — plainly (no template), or typeset into a journal's Word template with the
template's own styles, sections and headers, three-line tables, captions and references —
native-saved through **WPS Office** or Word, reopened, counted and rendered page by page for
review. Windows.

把 Markdown 稿件转成带 Word 原生公式的可编辑 DOCX：不给模板就按 Markdown 原样转换；给了期刊 Word 模板就按
模板排版。经 WPS 或 Word 原生保存、重开、计数并逐页渲染检查。需要 MathType 公式时，接着用 docx-omml-to-mathtype。

## Two inputs

| Markdown | template | result |
|---|---|---|
| ✓ | — | plain DOCX (bundled neutral styles: Times New Roman/宋体, black bold headings, grid tables) |
| ✓ | ✓ | typeset in the template, through the template's profile (found by file hash; made once if missing) |

## Design

- **No conventions in code.** How Markdown is recognised (front-matter labels, caption words,
  note prefixes, reference headings, citation habits) and laid out (sizes, sections, table rules)
  is data: `scripts/default_rules.json` ← journal profile ← per-paper overrides.
- **Uncertainty is settled by scripts.** `probe_template` (facts of any template),
  `validate_profile` (the profile's decisions checked against the template), `probe_host` (the
  COM host proven on a copy of the document), `check_runtime`, and a build that refuses unmapped
  lines and lists every unconvertible formula.
- **Evidence.** Content ledger, formula manifest, preserved-part hashes, save-check metrics,
  page PNGs.

## Quick start

```powershell
pwsh -File scripts\check_runtime.ps1
pwsh -File scripts\make_docx.ps1 -ManuscriptPath paper.md -OutputDirectory out -OutputStem paper                       # plain
pwsh -File scripts\make_docx.ps1 -ManuscriptPath paper.md -TemplatePath journal.doc -OutputDirectory out -OutputStem paper  # template
```

Exit 2 in template mode = no profile for that template yet: `references/new_journal_profile.md`.
Requirements: `references/requirements.md`.

## Notes

- A profile contains the journal's template; check its terms before redistributing one.
- Inputs and profiles are never written; outputs are replaced only with `-Force`.
