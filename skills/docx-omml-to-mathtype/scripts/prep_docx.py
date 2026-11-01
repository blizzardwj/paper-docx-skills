"""Replace OMML (Word-native) equations in a .docx with text markers and export each as MathML
tuned for MathType's API.

usage: prep_docx.py SRC.docx OUT_MARKED.docx OUT_EQS.json [--only 1,7] [--xsl PATH]

Markers look like QQMTEQ007QQ; numbers follow document order of m:oMath in word/document.xml.
Every MathML transformation below works around a specific MathType failure; see
references/pitfalls.md before "simplifying" any of them.
"""
import argparse
import glob
import json
import re
import sys
import unicodedata
import zipfile

from lxml import etree

XSL_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\OMML2MML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\OMML2MML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\OMML2MML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\Office16\OMML2MML.XSL",
]
EQ_RE = re.compile(r"<m:oMathPara>.*?</m:oMathPara>|<m:oMath>.*?</m:oMath>", re.S)
OMATH_RE = re.compile(r"<m:oMath>.*?</m:oMath>", re.S)
BM_RE = re.compile(r"<w:bookmark(?:Start|End)\b[^>]*/>")
BOLD_TAG = "\ue001"                      # private-use chars marking bold / bold-italic runs through the XSLT
BOLDITALIC_TAG = "\ue002"
VARIANT_OF_TAG = {BOLD_TAG: "bold", BOLDITALIC_TAG: "bold-italic"}
SPACES = "\u00a0\u2000-\u200a\u202f\u205f\u3000"   # NBSP, en/em quads and spaces, thin, ideographic
SPACE_TOKEN = rf'<(?:mi|mo|mtext)(?:\s[^>]*)?>[{SPACES}]+</(?:mi|mo|mtext)>'
CJK = "\u4e00-\u9fff"
LETTERLIKE = {"\u212c": ("script", "B"), "\u2130": ("script", "E"), "\u2131": ("script", "F"),
              "\u210b": ("script", "H"), "\u2110": ("script", "I"), "\u2112": ("script", "L"),
              "\u2133": ("script", "M"), "\u211b": ("script", "R"), "\u212f": ("script", "e"),
              "\u210a": ("script", "g"), "\u2134": ("script", "o"),
              "\u2102": ("double-struck", "C"), "\u210d": ("double-struck", "H"),
              "\u2115": ("double-struck", "N"), "\u2119": ("double-struck", "P"),
              "\u211a": ("double-struck", "Q"), "\u211d": ("double-struck", "R"),
              "\u2124": ("double-struck", "Z"), "\u212d": ("fraktur", "C"), "\u210c": ("fraktur", "H"),
              "\u2111": ("fraktur", "I"), "\u211c": ("fraktur", "R"), "\u2128": ("fraktur", "Z")}
# "MATHEMATICAL <STYLE> CAPITAL A" -> MathML mathvariant
STYLE_VARIANT = {"BOLD": "bold", "ITALIC": "italic", "BOLD ITALIC": "bold-italic", "SCRIPT": "script",
                 "BOLD SCRIPT": "bold-script", "FRAKTUR": "fraktur", "BOLD FRAKTUR": "bold-fraktur",
                 "DOUBLE-STRUCK": "double-struck", "SANS-SERIF": "sans-serif",
                 "SANS-SERIF BOLD": "bold-sans-serif", "SANS-SERIF ITALIC": "sans-serif-italic",
                 "SANS-SERIF BOLD ITALIC": "sans-serif-bold-italic", "MONOSPACE": "monospace"}
FUNCTION_NAMES = {"arccos", "arcsin", "arctan", "arg", "cos", "cosh", "cot", "coth", "csc", "deg", "det",
                  "dim", "exp", "gcd", "hom", "inf", "ker", "lg", "lim", "liminf", "limsup", "ln", "log",
                  "max", "min", "Pr", "sec", "sin", "sinh", "sup", "tan", "tanh", "argmax", "argmin"}


def find_xsl(explicit):
    if explicit:
        return explicit
    for p in XSL_CANDIDATES + glob.glob(r"C:\Program Files*\Microsoft Office*\**\OMML2MML.XSL", recursive=True):
        try:
            open(p).close()
            return p
        except OSError:
            continue
    sys.exit("OMML2MML.XSL not found (ships with Microsoft Office); pass --xsl PATH")


def styled_letter(ch):
    """Mathematical alphanumeric (U+1D400..1D7FF) or letterlike symbol -> (mathvariant, plain char),
    else None. 𝒳 -> (script, X), 𝔼 -> (double-struck, E), ℝ -> (double-struck, R), 𝛂 -> (bold, α)."""
    if ch in LETTERLIKE:
        return LETTERLIKE[ch]
    if not 0x1D400 <= ord(ch) <= 0x1D7FF:
        return None
    m = re.match(r"MATHEMATICAL (.+?) (CAPITAL|SMALL|DIGIT) (.+)$", unicodedata.name(ch, ""))
    if not m or m.group(1) not in STYLE_VARIANT:
        return None
    style, kind, base = m.groups()
    candidates = ([f"DIGIT {base}"] if kind == "DIGIT" else
                  [f"LATIN {kind} LETTER {base}", f"GREEK {kind} LETTER {base}", f"GREEK {base}"])
    for name in candidates:
        try:
            return STYLE_VARIANT[style], unicodedata.lookup(name)
        except KeyError:
            continue
    return None


def normalize_for_mathtype(mml):
    # math alphabets (script, double-struck, fraktur, bold…): MathType has no glyphs for the astral
    # math alphanumerics (renders a box or "?") -> plain letter + mathvariant
    def styled(m):
        tag, ch = m.group(1), m.group(3)
        hit = styled_letter(ch)
        if not hit:
            return m.group(0)
        tag = "mn" if hit[1].isdigit() else tag
        return f'<{tag} mathvariant="{hit[0]}">{hit[1]}</{tag}>'
    mml = re.sub(r"<(mi|mn|mo)(\s[^>]*)?>(.)</\1>", styled, mml)
    # \bar: the XSLT gives U+203E (overline) as <mi>, which MathType draws as a box; a macron <mo> is its bar
    mml = re.sub(r"<(?:mi|mo)(?:\s[^>]*)?>‾</(?:mi|mo)>", "<mo>¯</mo>", mml)
    # function names arrive as one upright <mi> per letter ("l","o","g") -> MathType then sets no
    # function space ("logp"); one <mi>log</mi> is recognised as a function
    def func(m):
        word = "".join(re.findall(r">([A-Za-z])<", m.group(0)))
        return f"<mi>{word}</mi>" if word in FUNCTION_NAMES else m.group(0)
    mml = re.sub(r'(?:<mi mathvariant="normal">[A-Za-z]</mi>){2,}', func, mml)
    # space tokens (\qquad, \text{ ...}) -> space characters in <mtext>; <mspace> crashes the API.
    # A space right before a CJK run is the word space inside \text{ ...}: keep it thin.
    # Space tokens come as <mi> (Office XSLT) or <mo> (pandoc/texmath, e.g. U+2001 for \qquad).
    mml = re.sub(rf'{SPACE_TOKEN}(?=<mi(?:\s[^>]*)?>[{CJK}])', "<mtext>\u2009</mtext>", mml)
    mml = re.sub(SPACE_TOKEN, "<mtext>\u2003</mtext>", mml)
    # cases: an open "{" fence around a table must be a stretchy fence (mtable may carry attributes)
    mml = re.sub(r'<mo>\{</mo><mrow>(<mtable\b[^>]*>.*?</mtable>)</mrow>(?:<mo\s*/>|<mo></mo>)',
                 r'<mfenced open="{" close="">\1</mfenced>', mml, flags=re.S)
    # runs of single CJK chars -> one <mtext> (typeset in the prefs' Text font, e.g. SimSun)
    mml = re.sub(rf'(?:<mi(?:\s[^>]*)?>[{CJK}]</mi>)+',
                 lambda m: "<mtext>" + "".join(re.findall(rf">([{CJK}])<", m.group(0))) + "</mtext>", mml)
    return mml


def make_converter(xsl_path, ns_decls):
    xslt = etree.XSLT(etree.parse(xsl_path))

    def to_mathml(omath_xml):
        # .NET serialisers may write xml:space as e.g. d3p1:space with xmlns:d3p1 bound to the XML
        # namespace, which lxml rejects; rewrite such prefixes back to xml:
        for pfx in re.findall(r'xmlns:(\w+)="http://www\.w3\.org/XML/1998/namespace"', omath_xml):
            omath_xml = re.sub(rf'\s+xmlns:{pfx}="http://www\.w3\.org/XML/1998/namespace"', "", omath_xml)
            omath_xml = re.sub(rf'\b{pfx}:(\w+)=', r"xml:\1=", omath_xml)
        wrapper = etree.fromstring(f"<wrap {ns_decls}>{omath_xml}</wrap>".encode("utf-8"))
        # OMML2MML.XSL loses bold: m:sty="b"/"bi" drops Latin letters entirely (an empty <mi/>) and
        # never emits mathvariant bold. Tag each bold char, make the run plain (b) / italic (bi) for
        # the XSLT, and restore the variant afterwards.
        m_ns = wrapper.nsmap["m"]
        for sty in wrapper.iter(f"{{{m_ns}}}sty"):
            val = sty.get(f"{{{m_ns}}}val")
            if val in ("b", "bi"):
                tag = BOLD_TAG if val == "b" else BOLDITALIC_TAG
                for t in sty.getparent().getparent().iter(f"{{{m_ns}}}t"):
                    t.text = "".join(tag + ch for ch in (t.text or ""))
                sty.set(f"{{{m_ns}}}val", "p" if val == "b" else "i")
        mml = str(xslt(etree.ElementTree(wrapper[0])))
        mml = re.sub(r"<\?xml[^>]*\?>", "", mml).strip()
        mml = mml.replace("mml:", "").replace("xmlns:mml=", "xmlns=")

        # the XSLT usually splits the tag char into its own <mi>; move the variant onto the next token
        for tag, variant in VARIANT_OF_TAG.items():
            mml = re.sub(rf'<mi(?:\s[^>]*)?>{tag}</mi>\s*<(mi|mn|mo)(?:\s[^>]*)?>([^<]*)</\1>',
                         rf'<\1 mathvariant="{variant}">\2</\1>', mml)

        def bold(m):
            tag, attrs, body = m.group(1), m.group(2) or "", m.group(3)
            hits = [t for t in VARIANT_OF_TAG if t in body]
            if not hits:
                return m.group(0)
            attrs = re.sub(r'\s*mathvariant="[^"]*"', "", attrs)
            for t in VARIANT_OF_TAG:
                body = body.replace(t, "")
            return f'<{tag}{attrs} mathvariant="{VARIANT_OF_TAG[hits[0]]}">{body}</{tag}>'
        mml = re.sub(r"<(mi|mn|mo|mtext)(\s[^>]*)?>([^<]*)</\1>", bold, mml)
        assert not any(t in mml for t in VARIANT_OF_TAG), "bold tag left in MathML"

        mml = normalize_for_mathtype(mml)
        # inline-mode MathML crashes MathType's API -> always block mode (the object is still
        # placed inline in the text); the clipboard text is read through the ANSI code page ->
        # every non-ASCII char as a numeric entity
        mml = mml.replace("<math ", '<math display="block" ', 1)
        return "".join(ch if ord(ch) < 128 else f"&#x{ord(ch):X};" for ch in mml)

    return to_mathml


def marker(i):
    return f"QQMTEQ{i:03d}QQ"


PPR_AFTER_JC = r"<w:(?:textDirection|textAlignment|textboxTightWrap|outlineLvl|divId|cnfStyle|rPr|sectPr|pPrChange)\b|</w:pPr>"


def center_display_paragraphs(xml, displays):
    """A display equation (m:oMathPara) is centred by default although its paragraph has no w:jc
    (pandoc output); the OLE object that replaces it is then left-aligned. Give such paragraphs
    w:jc=center unless the equation asked for left/right."""
    for e in displays:
        if e.get("math_jc") in ("left", "right"):
            continue
        at = xml.find(e["marker"])
        start = max(xml.rfind("<w:p>", 0, at), xml.rfind("<w:p ", 0, at))
        open_end = xml.index(">", start) + 1
        end = xml.index("</w:p>", at)
        para = xml[start:end]
        ppr = re.search(r"<w:pPr>.*?</w:pPr>|<w:pPr/>", para, re.S)
        if ppr and "<w:jc " in ppr.group(0):
            continue
        if ppr is None or ppr.group(0) == "<w:pPr/>":
            new_ppr = '<w:pPr><w:jc w:val="center"/></w:pPr>'
            para = (para[:ppr.start()] + new_ppr + para[ppr.end():]) if ppr else \
                para[:open_end - start] + new_ppr + para[open_end - start:]
        else:
            body = ppr.group(0)
            cut = re.search(PPR_AFTER_JC, body).start()
            para = para[:ppr.start()] + body[:cut] + '<w:jc w:val="center"/>' + body[cut:] + para[ppr.end():]
        xml = xml[:start] + para + xml[end:]
    return xml


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src"); ap.add_argument("out_docx"); ap.add_argument("out_json")
    ap.add_argument("--only", help="comma-separated equation numbers to convert (default: all)")
    ap.add_argument("--xsl", help="path to OMML2MML.XSL")
    args = ap.parse_args()
    only = {int(v) for v in args.only.split(",")} if args.only else None

    zin = zipfile.ZipFile(args.src)
    xml = zin.read("word/document.xml").decode("utf-8")
    for part in zin.namelist():
        if part != "word/document.xml" and part.endswith(".xml") and part.startswith("word/"):
            n = zin.read(part).decode("utf-8", "ignore").count("<m:oMath>")
            if n:
                print(f"WARNING: {part} has {n} equation(s); only word/document.xml is converted")
    root_tag = re.search(r"<w:document\b[^>]*>", xml).group(0)
    to_mathml = make_converter(find_xsl(args.xsl), " ".join(re.findall(r'xmlns:\w+="[^"]*"', root_tag)))

    eqs = []

    def replace(m):
        frag = m.group(0)
        omaths = OMATH_RE.findall(frag)
        if len(omaths) != 1:
            sys.exit(f"an m:oMathPara holds {len(omaths)} m:oMath; split it in WPS/Word first")
        i = len(eqs) + 1
        jc = re.search(r'<m:oMathParaPr>.*?<m:jc m:val="(\w+)"', frag, re.S)
        eqs.append({"index": i, "display": frag.startswith("<m:oMathPara>"), "marker": marker(i),
                    "math_jc": jc.group(1) if jc else None,
                    "text": "".join(re.findall(r"<m:t[^>]*>([^<]*)</m:t>", frag)),
                    "mathml": to_mathml(omaths[0]),
                    "converted": only is None or i in only})
        if only is not None and i not in only:
            return frag
        bms = BM_RE.findall(frag)
        starts = "".join(b for b in bms if "bookmarkStart" in b)
        ends = "".join(b for b in bms if "bookmarkEnd" in b)
        return f"{starts}<w:r><w:t>{marker(i)}</w:t></w:r>{ends}"

    new_xml = EQ_RE.sub(replace, xml)
    new_xml = center_display_paragraphs(new_xml, [e for e in eqs if e["display"] and e["converted"]])

    with zipfile.ZipFile(args.out_docx, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_xml.encode("utf-8")
            zout.writestr(item, data)
    with open(args.out_json, "w", encoding="utf-8") as f:
        json.dump(eqs, f, ensure_ascii=False, indent=1)
    print(f"equations: {len(eqs)} (display {sum(e['display'] for e in eqs)}), "
          f"to convert: {sum(e['converted'] for e in eqs)}")


if __name__ == "__main__":
    main()
