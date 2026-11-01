# Pitfalls and fixes

Problems met while building and porting this pipeline, each with its cause and where the fix
lives. Look a symptom up here before experimenting. Journal-specific quirks belong in the
profile's `profile.md`, not here.

**Adding an entry** (this file is the error log — SKILL.md stays an outline):
`N. **Symptom.** Cause. → Fix (file).` in the right section, plus one line in `CHANGELOG.md`.

## Template / profile

1. **Wrong font size although the style was applied.** Templates mix style sizes with direct run
   formatting, and the two often disagree. → `probe_template` shows style size and run sizes side
   by side; the mapping stores explicit `target_size_pt`; the builder writes explicit run sizes.
2. **Sample formatting contradicts the template's written instruction** (text says one size, the
   sample run has another). → Written requirement wins; record the conflict in `profile.md`
   (`probe_template` lists written requirements).
3. **Another journal's name/volume in headers, `XX` placeholders in footers, sample authors and
   photos** — templates are reused. → `header_fixes` / `footer_fixes`; unreferenced template media
   and their relationships are removed so they don't stay hidden in the package.
4. **Header text split over several runs** (common after WPS/Word conversion). → `replace_span`
   edits the runs in place and keeps PAGE fields; `validate_profile` checks it matches once.
5. **Caption/note paragraphs have no style of their own.** → Map roles to sample paragraphs, not
   to style names; the builder clones paragraph properties from the mapped sample.
6. **Template metadata (cached page count, "saved by …") taken as evidence.** → Render pages and
   approve them before `authoring_gate = passed`.
7. **A role's locator points at the wrong paragraph after the template changed.** →
   `expected_text_pattern` on every role; `validate_profile` fails on mismatch.
8. **Template has no inline picture to clone.** → Built-in minimal inline picture is used
   (`build_docx.ps1`); `validate_profile` warns.

## Markdown → DOCX build

9. **`Unmapped source lines` / `matches no rules.front_matter role` / `Missing required role`.**
   The manuscript doesn't follow the rules in force. → Express this paper's conventions as
   `paper_overrides.json` rules (rules.md) rather than rewriting the manuscript or the profile.
10. **Conventions hard-coded in scripts** (front-matter labels, caption words, note prefixes,
    reference heading, citation habits, per-table column widths, section count). They made the
    builder work for one paper only. → All moved to `default_rules.json` / profile / paper config;
    nothing journal- or paper-specific may be added to scripts.
11. **Duplicate numbering** (number in the heading text plus template auto-numbering). → Numbers
    come only from the source text; `numPr` is suppressed on generated paragraphs.
12. **Full-width figure/table/equation inside a multi-column body.** → Continuous section breaks
    cloned from the template's body section; first-page settings removed on added sections
    (`layout_config.json full_width_*`).
13. **Short tables broken across columns/pages.** → `keepNext` on all but the last row,
    `cantSplit` rows; `finalize_native` reports `spans_pages` and `make_docx` fails on any.
14. **Empty ZIP entries vanish.** PowerShell turns an `if`-expression yielding an empty byte
    array into `$null`. → Assign inside the branch (package writer).
15. **Variables silently overwritten** (namespace URI replaced by a formula; a rule object
    replaced by a paragraph). PowerShell names are case-insensitive (`$m` = `$M`,
    `$classification` = `$Classification`). → Distinct names for script-level values; never reuse
    `$W $M $R $PR $CT`. A second `-match` overwrites `$Matches` — copy groups first.
16. **Builds compared by file hash always differ.** New ZIP entries carry the current time. →
    Compare part by part; preserved template parts are SHA-checked by the builder.
17. **A revision-folder rule blocked a build.** → Only enforced with
    `outputs_must_stay_in_config_dir: true`.

## Formula engine

18. **Strict legacy engine (`cs`) rejects or degrades common LaTeX** (one environment, Σ limits
    always beside, `\qquad` nearly invisible, uppercase Greek italic, ASCII minus). → pandoc/texmath
    is the default; `cs` kept to reproduce old builds.
19. **latex2mathml + Office `MML2OMML.XSL`** dropped parentheses in superscripts and `<` in
    subscripts. → Rejected.
20. **pandoc OMML renders in the body font / italic digits in WPS.** texmath emits no run fonts. →
    `tex2omml.py` adds Cambria Math + East-Asian font + size to every `m:r`; the builder rescales
    the size per role.
21. **Odd gaps around n-ary operators; missing space inside `\text{ …}`.** texmath inserts U+200B
    and math zones drop plain spaces. → U+200B stripped; leading space in text runs → NBSP.
22. **Equations invisible to downstream tools.** Imported fragments carried `xmlns` on
    `<m:oMath>`. → Redundant namespace attributes removed after import.
23. **Malformed LaTeX printed as raw text** (pandoc only warns). → One marker paragraph per
    formula; failures are errors; all failing formulas are listed at once with line numbers.
24. **Slow builds with many formulas.** → One pandoc run for all formulas (prefetch).

## Host (WPS / Word)

25. **"Registered" ≠ "works".** An unlicensed Word still automated; a registered host may fail.
    → `probe_host.ps1` tests every host capability on a copy of the document; auto mode falls
    back to the next host.
26. **WPS: endless story-range chain.** `NextStoryRange` keeps returning the same range. → Stop
    on repeat (`finalize_native.ps1`, `probe_host.ps1`).
27. **WPS lacks `CheckCompatibility` / maybe `AutomationSecurity`; reports DOCX SaveFormat 12.** →
    Optional, warned; 12 and 16 accepted.
28. **DOC copy has pictures instead of equations** (both hosts). → DOCX is the master; for MathType,
    convert the DOCX first and save the DOC afterwards.
29. **Hosts rewrite OMML/section XML on save** (Unicode minus → ASCII, script letters → letter +
    `m:scr`, omitted `cols num="1"`). → Verify saved files structurally, not by bytes.
30. **COM attaches to a running instance holding the user's documents.** → Refuse when
    `Documents.Count > 0`; quit only owned instances; kill only `/Automation` processes started by
    the run.

## Layout

31. **Display equation wider than its column.** → `full_width_equations` in `layout_config.json`,
    or split the equation in the Markdown (content change — ask).

## Plain mode / template lookup

32. **`$…$` silently left as text.** pandoc does not treat `$\frac{a$` (unbalanced braces) or
    `$ x$` (space after the opening `$`) as math and gives no warning; raw LaTeX outside math is
    dropped from DOCX without a word. → `plain_docx.py` walks the pandoc AST and fails on any text
    `$` and any raw TeX (`--allow-literal-dollar` / `-AllowLiteralDollar` for real dollar signs).
33. **Missing picture only warned.** pandoc replaces it with its alt text. → Failure in
    `plain_docx.py`.
34. **Plain output differs from the Markdown preview.** pandoc turns a lone picture's alt text into
    a caption (duplicate of the manuscript's own caption line) and its default table style has no
    lines. → `-implicit_figures`; grid borders and keep-together rows added (`plain_docx.py`).
    pandoc's default look (blue Calibri/Aptos headings via theme fonts and colours) → bundled
    style reference without theme references (`make_plain_reference.py`).
36. **Chinese first-line indent leaks into table cells, pictures and display equations** (they
    inherit Body Text; an indented full-width picture runs past the right margin). → Compact style
    without indent; picture/equation paragraphs reset to no indent, pictures and tables centred,
    space after each table (`plain_docx.py`). New style element in `w:style` → keep the schema
    order (pPr, rPr before tblPr/tblStylePr) or hosts may ignore it.
35. **Same template, different file → profile not found.** The lookup is by SHA-256; a `.doc`
    and the `.docx` saved from it differ, and so does any re-save. → A profile records both
    (`template_visual_review.original_doc_sha256`, `format_mapping.reference_sha256`); an exit 2 on a
    template you believe is known: compare the printed hash, then add it or pass `-ProfileName`.
