# Changelog

Format: one line per change or correction; details of problems live in `references/pitfalls.md`
(referenced as #N).

## 0.3.0 — 2026-11-01

- Fixed on a new generic test sample (textbook formulas) and an edge-case document, found by
  reviewing the rendered page — all automated checks had passed:
  - bold Latin letters vanished (`\mathbf{w}`, OMML `m:sty="b"/"bi"`); bold-italic kept (#8);
  - `\mathbb`, `\mathfrak` and all other astral math alphabets → plain letter + `mathvariant` (#13);
  - `\bar` accent drawn as ◆ (#13a);
  - function names without function spacing (`logp`) (#9);
  - display equations from pandoc output left-aligned (#29a).
- Test case `converts-mixed-omml-equations` now uses generic formulas (same structure: 3 display +
  23 inline, 26 bookmarks in the same kinds of positions). Verified: old sample, new sample and
  edge document (14 equations) all pass every check; pages reviewed.

## 0.2.1 — 2026-10-31

- `prep_docx.py` accepts OMML whose `xml:space` was serialised under another prefix (#9a).
- Handles OMML produced by pandoc/texmath: all Unicode space characters in `mi`/`mo`/`mtext` (#14a),
  `cases` tables with attributes (#16). Verified: 52 equations (26 from each engine) all converted.
- `convert.ps1 -Only` now says that the "no OMML left" check is expected to fail on a partial run.

## 0.2.0 — 2026-10-31

- Restructured: SKILL.md is an outline; runtime details moved to `references/requirements.md`,
  problems and fixes to `references/pitfalls.md`, history here.
- Added `scripts/check_runtime.ps1` (preflight with install hints); `convert.ps1` runs it first.
- Found and documented a hidden dependency: `OMML2MML.XSL` ships only with Microsoft Office, not
  WPS (#3); added `-Xsl` option.
- Font check now resolves localized names (SimSun = 宋体) (#5).
- MathType cleanup kills only the `-mtrpc` API server's own `MathTypeLib.exe`, not a user's editor (#22).
- `verify_docx.py` compares `w:t` text with whitespace collapsed; it falsely failed when WPS
  normalised spaces (#28). Re-verified: manuscript and eval sample both pass all checks.
- Added test case `evals/converts-mixed-omml-equations` (anonymised: equation paragraphs only).
- Correction: display-equation count was double-counted via `<m:oMathParaPr>` (#7); the sample
  manuscript has 3 display + 23 inline equations, not 6 + 20.

## 0.1.0 — 2026-10-30

- First version, extracted from a manual WPS + MathType 7 session on a 26-equation manuscript.
- Pipeline: OMML → markers + MathML (Office XSLT) → MathType API via clipboard, one validated
  worker process per equation → WPS COM `Selection.Paste` → restore bookmarks → verify → PDF/PNG.
- Workarounds for #1, #2, #4, #6, #8–#27, #29.
- Verified end to end: 26/26 equations converted on the first worker, all checks passed.
