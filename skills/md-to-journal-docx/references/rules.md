# Rules: how the Markdown is recognised and laid out

All conventions are data. Three layers, later wins:

1. `scripts/default_rules.json` — neutral defaults (English and Chinese words both recognised).
2. `profiles/<journal>/profile.json` → `"rules"` — the journal's conventions.
3. `<paper config>/paper_overrides.json` → `"rules"` — one paper's deviations.

Dictionaries merge key by key; arrays and scalars replace. Check the merged result against the
template with `validate_profile.ps1 -ProfileDir … [-PaperOverridesPath …]`.

## markdown

| key | default | meaning |
|---|---|---|
| `body_start_pattern` | `^##\s+\S` | first line matching it starts the body; everything above is front matter |
| `heading_roles` | `{"1": section_heading, "2": subsection_heading, "3": third_level_heading}` | `##`=1, `###`=2, … → mapping role; deeper levels use the deepest role listed |
| `references_heading_pattern` | `^(References|Bibliography|参考文献)$` | heading text that starts the reference list (`[n] …` lines) |
| `appendix_heading_pattern` | `^(Appendix|附录)` | heading that ends reference mode |
| `figure_words`, `table_words` | `Figure, Fig., 图` / `Table, 表` | caption lines are `**<word> N …** optional note`; image alt text `<word> N` ties a picture to figure N |
| `table_note_pattern` | `^(Notes?|注)[：:.]` | paragraph right after a table that is its note |
| `citation_pattern` | `\[(?:\d+(?:\s*[-–,，]\s*\d+)*)\]` | in-text citation token |
| `citation_superscript` | `true` | citations as superscript |
| `citation_baseline_after_pattern` | `null` | if the text before a citation matches, keep it on the baseline (e.g. "see [3]") |
| `italic_token_pattern` | `null` | extra token rendered italic (e.g. a statistic symbol); usually a paper-level setting |

Always recognised: `$…$` inline math, `$$…$$` display math (may span lines), `![alt](path)`
pictures, pipe tables with a `---` separator row (`---:` = right-aligned), `[n] …` reference
entries, `1. …` / `- …` list items, `**bold**`, `*italic*`, `` `code` ``. Fenced code blocks are
rejected (no template mapping). Every non-empty line must be recognised or the build stops with
the unmapped line numbers.

## front_matter

| key | meaning |
|---|---|
| `roles` | ordered list. A role with `match` (regex, group 1 = content) claims the first line it matches, tried in list order; a role with `unlabelled: n` takes the n-th front-matter line that matched nothing. `label` = text printed bold before the content (`label_east_asia_font` optional); `required` = build fails if absent |
| `render_order` | roles in output order; roles not in the Markdown print `missing_text` |
| `render.<role>` | `align` (left/center/both/right), `keep_next`, `line_spacing`, `suppress_hyphenation` |
| `missing_text` | placeholder for absent roles (default `TBD`) |
| `label_separator` | between label and content (default space) |
| `document_title_role` | role used for document properties and `title_break_before` |

Every role used here must exist in `format_mapping.json` → `role_mapping` (validated).

## body_roles

Mapping role names for `body`, `references_heading`, `reference_entry`, `figure_caption`,
`table_caption`, and the optional secondary-language captions `figure_caption_secondary` /
`table_caption_secondary` (printed only for numbers listed in `english_captions.json`).

## captions, tables, paragraphs, references

| key | default |
|---|---|
| `captions.secondary_format` | `{"figures": "Figure {n}  {title}", "tables": "Table {n}  {title}"}` |
| `captions.note_size_pt` | 8 (caption notes and table notes) |
| `captions.line_spacing` | 1.15 (overridable per paper via `layout_config.caption_line_spacing_multiplier`) |
| `tables.font_size_pt` | `null` = `format_mapping.tables.body_size_pt` |
| `tables.cell_padding_top_bottom_twips` | 20 |
| `tables.cell_side_margin_twips` | `{"column": 30, "full": 65}` |
| `tables.rule_sizes_eighth_pt` | `{"outer": 12, "header": 6}` (three-line table) |
| `paragraphs` | `body_first_line_chars` 2, `line_spacing` 1.15, `equation_line_spacing` 1, `reference_line_spacing` 1 |
| `references.hanging_indent_twips` | 300 |

## sections

`{"front": i, "body": j, "tail": [k, …]}` — indices of the template's `w:sectPr` (document
order). Front matter ends with section `front`; the body uses section `body` (its column count
is the normal layout; `full`-width objects switch to one column with continuous breaks);
`tail` sections close the document in order, the last becoming the final section.
`author_info_placeholder` (optional text) goes into the first tail section. No default — the
profile must declare them; `validate_profile.ps1` checks them against the template.

**Tested layouts:** templates with separate front/body/tail sections (single-column front
matter, multi-column body). A single-section template (front = body = tail) is untested — run
`validate_profile.ps1`, build the profile's example and review pages before relying on it.
