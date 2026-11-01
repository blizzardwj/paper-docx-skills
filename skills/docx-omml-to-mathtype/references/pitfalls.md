# Pitfalls and fixes

Every problem met while building and running this pipeline, with its cause and the fix that is
now in the scripts. Look a symptom up here before experimenting; read the relevant entry before
changing a script, because most odd-looking lines exist because of one of these.

**Adding an entry** (this file is the error log — SKILL.md stays an outline):
`N. **Symptom.** Cause. → Fix (file).` under the right section, plus one line in `CHANGELOG.md`.

## Environment

1. **MathType "Convert Equations" says "install VBA first" in WPS.** WPS personal edition has no
   VBA and the MathType Office add-in is VBA. → Don't use the add-in; call MathType's C API
   (`MT6.dll`) directly (`MtApi.cs`).
2. **Word is installed but unlicensed.** Read-only mode; macros won't run. → Use WPS over COM.
3. **`OMML2MML.XSL` missing on WPS-only machines.** It ships only with Microsoft Office.
   → `check_runtime.ps1` reports it; `-Xsl <path>` accepts a copy (requirements.md).
4. **Project `.venv` broken ("uv trampoline failed").** → The skill keeps its own env in
   `%LOCALAPPDATA%\docx-omml-to-mathtype\venv` (`convert.ps1`).
5. **SimSun reported missing although installed.** Chinese Windows lists it as 宋体. → Check
   fonts with `New-Object System.Drawing.FontFamily <name>`, which resolves aliases
   (`check_runtime.ps1`).
6. API constants/signatures were taken from MathType's own VBA (`MathType Commands 2016.dotm`,
   read with `olevba`): `mtxfmCLIPBOARD=-2, mtxfmFILE=-4, mtxfmPICT=6, mtxfmTEXT=7, mtxfmGIF=9,
   mtxfmPREF_USER=3, mtOLE2_EQUATION=8`.

## Reading the document

7. **Display equations double-counted.** Matching the prefix `<m:oMathPara` also hits
   `<m:oMathParaPr>`. → Count `<m:oMathPara>` with the closing `>` (`prep_docx.py`).
8. **Bold lost (`\mathbf{1}`, OMML `m:sty="b"`) — and bold Latin letters vanish entirely**
   (`\mathbf{w}` → empty `<mi/>`; pandoc writes `m:sty="bi"`). OMML2MML.XSL never emits bold and
   drops Latin letters of `b`/`bi` runs; digits and Greek survive, which hid the letter case. →
   Tag each char (U+E001 bold, U+E002 bold-italic), set the run to `p`/`i` for the XSLT, then put
   `mathvariant="bold"`/`"bold-italic"` on the token; the XSLT splits the tag into its own `<mi>`
   (`prep_docx.py`). Seen only by looking at the rendered page: the width check passed because ⊤
   was still drawn.
9a. **`XMLSyntaxError: xml namespace URI mapped to wrong prefix`.** OMML written by a .NET
    serialiser carries `d3p1:space` with `xmlns:d3p1` bound to the XML namespace. → Rewrite such
    prefixes to `xml:` before parsing (`prep_docx.py`).
9. Sum limits beside ∑ are faithful to the original OMML (`limLoc=subSup`) — not a conversion
   error. **`logp` without a space**: the XSLT splits function names into upright `l`,`o`,`g`, so
   MathType sees no function. → Runs of upright single letters forming a known function name
   (log, det, max, sin, …) become one `<mi>log</mi>` (`prep_docx.py`).

## MathML handed to MathType

10. **MathML source typeset as literal text.** `MTXFormEqn(LOCAL, TEXT, …)` doesn't parse MathML
    from memory. → Put MathML on the clipboard (Unicode text plus the registered "MathML",
    "MathML Presentation", "application/mathml+xml" formats) and convert `CLIPBOARD, TEXT`
    (`mt_worker.ps1`).
11. **Process crash (0xC0000005) on inline equations.** MathML without `display`, or with
    `display="inline"`, crashes the API. → Always `display="block"`; the object is still placed
    inline (`prep_docx.py`).
12. **Wrong characters / "not an equation" (-7).** The clipboard text is read through the ANSI code
    page (GBK). → Every non-ASCII char becomes `&#xHHHH;` (`prep_docx.py`).
13. **Script letters (𝒴, U+1D4B4…) shown as ◆; 𝔼 (`\mathbb`) as "?".** No glyphs for the astral
    math alphabet. → Every Mathematical Alphanumeric Symbol (U+1D400–1D7FF: script, double-struck,
    fraktur, bold, sans-serif, monospace …) and letterlike ℒ, ℝ, ℕ, ℭ … becomes the plain letter
    + `mathvariant` (`prep_docx.py`).
13a. **`\bar{X}` shows a ◆ accent.** The XSLT gives the accent as `<mi>`U+203E (overline). →
    `<mo>`U+00AF (macron), which MathType draws as a bar (`prep_docx.py`).
14. **Crash on `<mspace>`.** → Spaces become `<mtext>` with U+2003 (`\qquad`) or U+2009 (the word
    space before `\text{ 中文}`) (`prep_docx.py`).
14a. **`��` where a wide space should be.** OMML made by pandoc/texmath puts `\qquad` as U+2001 in
    an `<mo>`; only `<mi>` NBSP/en/em/thin were handled. → Any `<mi|mo|mtext>` holding only Unicode
    space characters (U+00A0, U+2000–U+200A, U+202F, U+205F, U+3000) becomes an em/thin `<mtext>`.
15. **Chinese inside an equation overlaps neighbours.** Text style was Times New Roman. → CJK runs
    merged into one `<mtext>`; prefs `Text=<CJK font>` (`prep_docx.py`, `convert.ps1 -CjkFont`).
16. **`cases` brace doesn't stretch.** The XSLT emits `<mo>{</mo><mrow><mtable…>…</mtable></mrow><mo/>`
    (the `mtable` carries `columnalign` when the OMML came from pandoc). → Rewritten to
    `<mfenced open="{" close="">` (`prep_docx.py`).

## MathType API stability

17. **Crash after `MTXFormSetPrefs`.** MathType keeps the prefs *pointer*; .NET frees its temporary
    ANSI copy. → Prefs held in `Marshal.StringToHGlobalAnsi` memory for the process lifetime (`MtApi.cs`).
18. **Random access violations, garbage `dims`.** The native code writes more than the VBA-declared
    structs/buffers hold. → 64 KB unmanaged src/dst buffers and a 4 KB dims block (`MtApi.cs`).
19. **Return code 0 but the result is the MathML source as text** (random). → The worker renders
    the clipboard copy to GIF and rejects it if wider than max(60, 25 px × visible chars); literal
    text is ~6 px per markup char (`mt_worker.ps1`).
20. **Random crashes of the calling process.** → One worker process per equation, up to 6 workers,
    90 s timeout each (`wps_replace.ps1`).
21. **`mtxfmPREVIOUS` (re-emit last equation) returns -9.** → Not used; the clipboard copy is
    validated instead (#19).
22. **MathType dialog "too many windows open", run hangs.** The `MathType.exe -mtrpc` API server
    keeps hidden windows from every conversion and crash. → Kill that server (and only its own
    `MathTypeLib.exe` child) after every paste and before each retry (`wps_replace.ps1`). Since
    this fix every equation has succeeded on the first worker.
23. The clipboard OLE data survives the worker's exit but not the server kill → paste first, then
    kill (`wps_replace.ps1`).

## WPS COM

24. **`Range.Paste()` throws "Unexpected HRESULT".** → `Range.Select(); Selection.Paste()`.
    `PasteSpecial` variants also fail (`wps_replace.ps1`).
25. **COM may attach to a running WPS.** → Record `Documents.Count` first, `Quit()` only if it was
    0, and only kill `/Automation` processes started by the run (`wps_replace.ps1`, `export_pdf.ps1`).
26. **Bookmarks around equations disappear.** Pasting over a marker that a bookmark encloses
    deletes the bookmark. → `restore_bookmarks.py` re-wraps the n-th object run with the n-th
    equation's original bookmarks.
27. WPS writes the standard `<w:object><v:shape…/><o:OLEObject ProgID="Equation.DSMT4"…/>` with
    `word/embeddings/oleObjectN.bin` — same as Word + MathType, so the result also opens in Word.

## Verification

28. **"body text unchanged" fails on whitespace only.** WPS re-saves without the newline after the
    XML declaration and collapses spaces in `w:t` runs that lack `xml:space="preserve"`.
    → Compare the concatenated `w:t` text with whitespace runs collapsed (`verify_docx.py`).

## Layout

29a. **Display equation left-aligned after conversion.** OMML centres an `m:oMathPara` by default,
    but pandoc's paragraph has no `w:jc`; the OLE object then sits at the left margin. → Such
    paragraphs get `w:jc=center` unless `m:oMathParaPr` says left/right (`prep_docx.py`).
29. **A long display equation overflows its column.** Already true in the original; MathType
    objects are fixed-size pictures and never auto-break like OMML. → Report it and offer to split
    the equation (content change — ask first).
