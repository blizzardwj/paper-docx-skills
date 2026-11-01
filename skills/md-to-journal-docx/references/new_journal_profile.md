# Making a profile for a new journal template

The builder trusts the profile, so the profile is made from evidence and checked by script.
The agent supplies judgement (which sample paragraph is which role, which written rule wins);
scripts supply facts and verification. Record every judgement so it can be reviewed.

Normally you arrive here because `make_docx.ps1 -TemplatePath …` (or `resolve_template.ps1`)
exited 2: no profile carries this template's SHA-256. It has already prepared `<work>/template/`:
`template.docx` (a `.doc`/`.dot`/`.dotx` saved as DOCX through the host; the original untouched),
`template_probe.json` / `.md`, `template.pdf` and `pages/pageNN.png`.

## 1. Create the profile folder

`profiles/<name>/` (short ASCII name). Copy `<work>/template/template.docx` →
`template_reference.docx`; never edit it afterwards. In `template_visual_review.json` record
`original_doc_sha256` = the SHA-256 of the file the user gave (printed by `resolve_template`), so
the next run with that file finds this profile; `format_mapping.reference_sha256` = the SHA-256 of
`template_reference.docx`.

## 2. Probe (facts)

Already done by `resolve_template` (to redo: `pwsh -File scripts/probe_template.ps1 -TemplatePath
profiles/<name>/template_reference.docx -OutDir <work>/probe`).

Read `template_probe.md`: sections (columns, size, margins), every body paragraph (locator,
style and its size, actual run sizes and fonts, bold, alignment, text), tables, header/footer
runs and fields, and **written requirements** — sample text that states formatting
("Title: 18 pt bold", "题目 黑体22点").

Look at the rendered pages (`<work>/template/pages/`) next to it.

## 3. Decide (agent) — write each decision down

| decision | how to decide | where it goes |
|---|---|---|
| role → sample paragraph | match the sample text/position to the role (title, authors, affiliations, abstract, keywords, headings by level, body, captions, notes, references heading/entry, acknowledgements); pick the paragraph whose formatting is the sample, not an instruction line | `format_mapping.role_mapping.<role>.locator` + `expected_text_pattern` (a regex on that sample text, so a wrong locator fails validation) |
| target font/size/bold per role | **written requirement > direct run formatting > style declaration**; when they disagree, write down all three | `target_*` fields; conflicts in `profile.md` |
| section roles | which section ends the front matter, which holds the body (its column count), which close the document | `profile.json rules.sections`, `format_mapping.sections` (one entry per template section, in order) |
| sample table | the table that shows the journal's table style | `format_mapping.tables.locator`, `body_size_pt` |
| sample text to neutralise | another journal's name/volume in headers, `XX` placeholders, sample authors/photos | `header_fixes` / `footer_fixes` |
| Markdown conventions | how papers for this journal are written (front-matter labels, caption words, references heading, citation habits) | `profile.json rules` (see rules.md) |
| anything not decidable from the template | ask the user | `format_mapping.unresolved` |

Sanity rules: never infer missing facts (author data, classification codes) — use placeholders;
a role's locator must point at a paragraph, not a section break; prefer the first, clean sample.

## 4. Validate (script)

`pwsh -File scripts/validate_profile.ps1 -ProfileDir profiles/<name>` → fix every FAIL, read
every WARN. It checks the template hash, approval gate, every locator and sample-text pattern,
section count/columns, that every role the rules use is mapped, regexes compile, header/footer
fixes match exactly once, and fonts.

## 5. Approve and test

1. Show the user the rendered template pages and the decision table; on approval set
   `authoring_gate: "passed"` and write `template_visual_review.json`.
2. Write `profiles/<name>/example/manuscript.md` (a short paper in this journal's conventions:
   front matter, headings, inline and display math, a table with a note, a figure, references)
   and run `make_docx.ps1` on it; review every page.
3. Summarise the conventions in `profiles/<name>/profile.md`; add any new generic problem to
   `pitfalls.md`.
