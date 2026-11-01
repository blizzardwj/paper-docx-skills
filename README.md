# paper-docx-skills

Claude Code skills for getting a paper written in **Markdown** into the **Word/WPS** form journals
ask for — on Windows, with **WPS Office** (no Microsoft Word licence, no VBA needed).

用 Claude Code 把 Markdown 论文变成期刊要的 Word/WPS 文档：可选按期刊模板排版，公式先生成 Word 原生公式，
需要时再批量转成 MathType。Windows，只需 WPS（不需要 Word 授权和 VBA）。

```
paper.md ──md-to-journal-docx──▶ paper.docx (Word-native equations) ──docx-omml-to-mathtype──▶ paper_mathtype.docx
           [+ journal template]                                          [MathType 7]
```

| skill | what it does |
|---|---|
| [`md-to-journal-docx`](skills/md-to-journal-docx) | Markdown → editable DOCX with Word-native (OMML) equations. **No template**: clean neutral styles (Times New Roman/宋体, grid tables). **With a journal template**: the template's styles, sections, headers, multi-column layout, three-line tables, captions, references — learned once per template as a checked *profile*. Native-saved through WPS or Word, counted and rendered page by page. |
| [`docx-omml-to-mathtype`](skills/docx-omml-to-mathtype) | Every Word-native equation in a DOCX → a real MathType object (`Equation.DSMT4`), through MathType's own API and WPS COM; bookmarks and text kept, each equation render-checked. |

Each skill is self-contained: `SKILL.md` (the outline Claude follows), `references/` (requirements,
pitfalls met and fixed, rules), `scripts/`, `evals/` (test cases).

## Requirements

- Windows 10/11 x64, PowerShell 7 (`winget install Microsoft.PowerShell`)
- WPS Office for Windows (personal edition is enough) — or desktop Word
- [uv](https://docs.astral.sh/uv/) or Python 3.10+ (the skills create their own environments:
  pandoc via `pypandoc_binary`, lxml, PyMuPDF)
- for `docx-omml-to-mathtype`: MathType 7 and `OMML2MML.XSL` from any Microsoft Office install

Each skill's `scripts/check_runtime.ps1` reports what is missing and how to install it; details in
`skills/*/references/requirements.md`.

## Install

As a Claude Code plugin:

```
/plugin marketplace add blizzardwj/paper-docx-skills
/plugin install paper-docx-skills@paper-docx-skills
```

Or copy the folders under `skills/` to `%USERPROFILE%\.claude\skills\`.

Then just ask, e.g. "把 paper.md 转成 Word", "按这个期刊模板 journal.doc 排版 paper.md",
"把 paper.docx 里的公式转成 MathType".

## Use without Claude

```powershell
# Markdown -> DOCX, plain or with a template
pwsh -File skills\md-to-journal-docx\scripts\make_docx.ps1 -ManuscriptPath paper.md -OutputDirectory out -OutputStem paper
pwsh -File skills\md-to-journal-docx\scripts\make_docx.ps1 -ManuscriptPath paper.md -TemplatePath journal.doc -OutputDirectory out -OutputStem paper
# Word-native equations -> MathType
pwsh -File skills\docx-omml-to-mathtype\scripts\convert.ps1 -InputDocx out\paper.docx
```

A template without a profile yet makes `make_docx` stop with exit code 2 after preparing the
evidence (probe, rendered pages); building the profile needs judgement — Claude does it with you
(`skills/md-to-journal-docx/references/new_journal_profile.md`).

## Safety

Inputs are never modified; outputs are replaced only with `-Force`. Host automation refuses to use
a WPS/Word instance that already holds your documents and only closes processes it started.

## Notices

Not affiliated with Wiris/Design Science (MathType), Kingsoft (WPS) or Microsoft; no MathType,
WPS or Microsoft code is included. Journal templates are not included — a profile you build contains
your journal's template, so check its terms before sharing it.

## License

[MIT](LICENSE)
