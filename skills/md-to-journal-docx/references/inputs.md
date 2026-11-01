# Inputs

| kind | given as | mode | lives in |
|---|---|---|---|
| Markdown manuscript | `-ManuscriptPath` (required) | both | the paper's repository |
| journal template | `-TemplatePath` (optional; absent = plain mode) | template | wherever the user keeps it |
| journal profile (`profiles/<name>/`) | found from the template's SHA-256 (or `-ProfileName` / `-ProfileDir`) | template | this skill |
| per-paper config (`-PaperConfigDir`) | optional | template | the paper's repository |

Plain mode reads pandoc Markdown as is: `$…$` / `$$…$$` math, pipe/grid tables, `![alt](path)`
pictures relative to the Markdown file, headings by `#` level. Nothing else is configured. The
output follows what a Markdown preview shows: tables with grid lines, centred and kept on one page
where they fit; pictures centred, without an alt-text caption; `lang` = zh-CN when the text has CJK
characters (East-Asian font of equations: 宋体).

Styles come from `assets/plain_reference.docx`, or `plain_reference_zh.docx` for Chinese text
(body paragraphs indented two characters; table cells, pictures and display equations not):
body Times New Roman / 宋体 12 pt, 1.25 lines; title 18 pt and headings 16 / 14 / 12 pt bold,
Times New Roman / 黑体, black; tables and captions 10.5 pt; code Consolas. Both files are made by
`scripts/make_plain_reference.py` from pandoc's own default reference — change a style there and
re-run it, never edit the DOCX by hand. `plain_docx.py --reference-doc` takes another style
reference for a one-off run.

In template mode, how the Markdown is recognised is defined by the rules (`references/rules.md`):
check the profile's rules first, then express a paper's deviations as overrides.

## Per-paper config (all optional)

Pass a folder with `-PaperConfigDir`; single files can be given with `-…Path` parameters instead.

| file | purpose | shape |
|---|---|---|
| `paper_overrides.json` | rule overrides for this paper; `title_break_before` (layout-only line break in the document title before this substring); `table_caption_paren_to_note` (`{tables:[n…], pattern}` moves regex group 2 of a table caption into its note) | `{"rules": {…}, "title_break_before": "…", "table_caption_paren_to_note": {…}}` |
| `layout_config.json` | full-width objects, figure moves, table geometry | `{"full_width_tables":[n], "full_width_figures":[n], "full_width_equations":[n], "caption_line_spacing_multiplier":1.15, "figure_placements":[{"number":n, "after_heading_pattern":"…", "after_paragraphs":k \| "after_paragraph_pattern":"…", "include_table_number":m}], "table_layouts":{"n":{"column_weights":[…], "cell_side_margin_twips":30, "soft_breaks":[{"row":r,"column":c,"after":"text"}]}}}` |
| `figure_map.json` | replace a Markdown picture with a prepared file; width mode | `{"figures":[{"source":"<path as in the Markdown>", "path":"<relative to this JSON>", "width_mode":"column"\|"full"}]}` |
| `english_captions.json` | secondary-language caption per figure/table number | `{"figures":[{"number":1,"title":"…"}], "tables":[…]}` |
| `references_formatted.json` | display form of reference entries; `source_exact` must equal the Markdown entry without `[n]` (checked) | `{"references":[{"number":n, "source_exact":"…", "english":"…", "chinese":"optional second line", "notes":[]}]}` |
| `reference_translations.json` | older bilingual pairs used when no formatted entry exists | `{"references":[{"number":n, "original_chinese":"…", "english":"…"}]}` |
| `source_numbering_correction.json` | recorded in the build report only | free-form |

Without `layout_config.json` every object stays in the body's column layout and tables get equal
column widths. Display equations are numbered in order for `full_width_equations`.

Defaults matter: a long display equation does not wrap inside a column in every host — if one
overflows (page review), list it in `full_width_equations` and rebuild.

## Journal profile

| file | content |
|---|---|
| `template_reference.docx` | the template as DOCX, never modified |
| `format_mapping.json` | `role_mapping` (role → template paragraph `locator`, `expected_text_pattern`, target fonts/size/bold, indent, line spacing), `sections` (geometry per template section, in order), `tables.locator`, `reference_sha256`, `authoring_gate`, `unresolved` |
| `profile.json` | `rules` (see rules.md), `header_fixes` / `footer_fixes`, `front_matter.classification` (optional generated line cloned from a template paragraph), `fonts_required` |
| `template_visual_review.json` | record of the approved template pages |
| `profile.md` | the journal's conventions in prose + audit notes |
| `example/` | a short manuscript in this profile's conventions (smoke test) |
| `tools/` | scripts that produced this profile's audit (historical reference) |

`header_fixes`: `{"part": "word/header2.xml", "replace_run_compact": "<run text, whitespace ignored>", "with": "…"}`
or `{"part": …, "replace_span": "<text across runs>", "span_offset": i, "span_length": n, "with": "…", "expect": "<result>"}`.
`footer_fixes`: `{"part": …, "replace": [[old, new]…], "exact": [[whole-run, new]…], "append_suffix": {"when": [run…], "suffix": "…"}}`.
Each must match the template (validated); runs are edited in place so fields survive.
