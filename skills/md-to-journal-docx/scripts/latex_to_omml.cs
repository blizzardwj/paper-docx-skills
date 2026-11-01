// Task-local, dependency-free LaTeX subset -> editable Word OMML.
// Unsupported or malformed input fails explicitly; it is never silently dropped.
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace ManuscriptTypesetting {
    public sealed class TexNode {
        public string Kind, Value, Style;
        public List<TexNode> Children = new List<TexNode>();
        public List<TexNode> Sub, Sup;
        public List<List<List<TexNode>>> Rows;
        public TexNode(string kind, string value = "", string style = "") { Kind = kind; Value = value; Style = style; }
    }

    public sealed class TexParser {
        readonly string source; int pos;
        static readonly Dictionary<string, string> Symbols = new Dictionary<string, string> {
            {"theta","θ"},{"alpha","α"},{"beta","β"},{"gamma","γ"},{"Delta","Δ"},{"delta","δ"},
            {"in","∈"},{"land","∧"},{"lor","∨"},{"mid","|"},{"Vert","∥"},{"vert","|"},
            {"subseteq","⊆"},{"cup","∪"},{"rightarrow","→"},{"le","≤"},{"ge","≥"},
            {"leq","≤"},{"geq","≥"},{"times","×"},{"cdot","·"},{"neq","≠"},{"pm","±"},
            {"dots","…"},{"ldots","…"},{"cdots","⋯"},{"infty","∞"},{"ell","ℓ"},
            {"log","log"},{"min","min"},{"max","max"},{"exp","exp"},{"Pr","Pr"}
        };
        public TexParser(string text) { source = text; }
        bool At(string token) { return source.IndexOf(token, pos, StringComparison.Ordinal) == pos; }
        void SkipSpaces() { while (pos < source.Length && Char.IsWhiteSpace(source[pos])) pos++; }
        Exception Error(string message) { return new FormatException(message + " at " + pos + " in " + source); }

        public List<TexNode> Parse() {
            var nodes = Sequence("", "");
            if (pos != source.Length) throw Error("Unconsumed LaTeX");
            return nodes;
        }
        List<TexNode> Sequence(string stop, string style) {
            var nodes = new List<TexNode>();
            while (pos < source.Length) {
                if (stop.Length > 0 && At(stop)) { pos += stop.Length; return nodes; }
                if (source[pos] == '}' || At(@"\right") || At(@"\end{")) throw Error("Unexpected group terminator");
                if (Char.IsWhiteSpace(source[pos])) { pos++; continue; }
                var node = Atom(style);
                if (node == null) continue;
                while (true) {
                    SkipSpaces();
                    if (pos >= source.Length || (source[pos] != '_' && source[pos] != '^')) break;
                    char marker = source[pos++];
                    if (marker == '_' && node.Sub != null || marker == '^' && node.Sup != null) throw Error("Duplicate math script");
                    if (marker == '_') node.Sub = Argument(style); else node.Sup = Argument(style);
                }
                nodes.Add(node);
            }
            if (stop.Length > 0) throw Error("Unclosed group " + stop);
            return nodes;
        }
        List<TexNode> Argument(string style) {
            SkipSpaces();
            if (pos >= source.Length) throw Error("Missing argument");
            if (source[pos] == '{') { pos++; return Sequence("}", style); }
            var atom = Atom(style);
            if (atom == null) throw Error("Empty math argument");
            return new List<TexNode> { atom };
        }
        string RawGroup() {
            SkipSpaces();
            if (pos >= source.Length || source[pos] != '{') throw Error("Expected braced text");
            int start = ++pos, depth = 1;
            while (pos < source.Length) {
                if (source[pos] == '{') depth++;
                if (source[pos] == '}' && --depth == 0) { string text = source.Substring(start, pos - start); pos++; return text; }
                pos++;
            }
            throw Error("Unclosed text group");
        }
        string Delimiter() {
            SkipSpaces();
            if (pos >= source.Length) throw Error("Missing delimiter");
            if (source[pos] == '\\') {
                pos++;
                if (At("lbrace")) { pos += 6; return "{"; }
                if (At("rbrace")) { pos += 6; return "}"; }
                if (At("Vert")) { pos += 4; return "∥"; }
                if (At("vert")) { pos += 4; return "|"; }
            }
            char c = source[pos++];
            return c == '.' ? "" : c.ToString();
        }
        TexNode Atom(string style) {
            if (pos >= source.Length) throw Error("Missing atom");
            char c = source[pos++];
            if (c == '{') { var g = new TexNode("group", "", style); g.Children = Sequence("}", style); return g; }
            if (c == '_' || c == '^') throw Error("Script has no base");
            if (c != '\\') return new TexNode("text", c.ToString(), style);
            if (pos >= source.Length) throw Error("Trailing backslash");
            int start = pos;
            if (Char.IsLetter(source[pos])) while (pos < source.Length && Char.IsLetter(source[pos])) pos++;
            else pos++;
            string command = source.Substring(start, pos - start);
            if (command == "left") {
                string begin = Delimiter(); var d = new TexNode("delimiter", begin, style);
                d.Children = Sequence(@"\right", style); d.Value += "\u0001" + Delimiter(); return d;
            }
            if (command == "begin") {
                string environment = RawGroup();
                if (environment != "cases") throw Error("Unsupported environment " + environment);
                string end = @"\end{cases}"; int endPos = source.IndexOf(end, pos, StringComparison.Ordinal);
                if (endPos < 0) throw Error("Unclosed cases");
                string body = source.Substring(pos, endPos - pos); pos = endPos + end.Length;
                var matrix = new TexNode("cases", "", style); matrix.Rows = new List<List<List<TexNode>>>();
                foreach (string row in body.Split(new string[] { @"\\" }, StringSplitOptions.None)) {
                    var cells = new List<List<TexNode>>();
                    foreach (string cell in row.Split('&')) cells.Add(new TexParser(cell).Parse());
                    if (cells.Count != 2) throw Error("Cases row must have two columns");
                    matrix.Rows.Add(cells);
                }
                return matrix;
            }
            if (command == "frac") { var f = new TexNode("fraction", "", style); f.Children = Argument(style); f.Sub = Argument(style); return f; }
            if (command == "sqrt") { var r = new TexNode("radical", "", style); r.Children = Argument(style); return r; }
            if (command == "hat" || command == "bar" || command == "tilde") {
                var a = new TexNode("accent", command == "hat" ? "̂" : command == "bar" ? "̅" : "̃", style); a.Children = Argument(style); return a;
            }
            if (command == "sum") return new TexNode("sum", "∑", style);
            if (command == "text") return new TexNode("text", RawGroup(), "plain");
            if (command == "mathcal" || command == "mathbf" || command == "mathrm" || command == "mathit" || command == "mathbin") {
                string newStyle = command == "mathcal" ? "script" : command == "mathbf" ? "bold" : command == "mathrm" ? "plain" : command == "mathit" ? "italic" : style;
                var g = new TexNode("group", "", newStyle); g.Children = Argument(newStyle); return g;
            }
            if (command == "!" ) return null;
            if (command == "," || command == ";" || command == ":" || command == "quad" || command == "qquad" || command == " ") return new TexNode("text", command == "qquad" ? "  " : " ", "plain");
            if (command == "{" || command == "}" || command == "[" || command == "]" || command == "(" || command == ")" || command == "|" || command == "_" || command == "%") return new TexNode("text", command, style);
            string symbol;
            if (Symbols.TryGetValue(command, out symbol)) return new TexNode("text", symbol, command == "log" || command == "min" || command == "max" || command == "exp" || command == "Pr" ? "plain" : style);
            throw Error("Unsupported LaTeX command \\" + command);
        }
    }

    public static class Omml {
        const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        static XmlElement E(XmlDocument doc, string name, string val = null) { var e = doc.CreateElement("m", name, M); if (val != null) e.SetAttribute("val", M, val); return e; }
        static XmlElement Container(XmlDocument doc, string name, List<TexNode> children, double pt) { var e = E(doc, name); Append(doc, e, children, pt); return e; }
        static XmlElement Script(XmlDocument doc, TexNode n, XmlElement atom, double pt) {
            // Fractions use Sub for their denominator, not a mathematical script.
            if (n.Kind == "fraction" || n.Sub == null && n.Sup == null) return atom;
            string kind = n.Sub != null && n.Sup != null ? "sSubSup" : n.Sub != null ? "sSub" : "sSup";
            var script = E(doc, kind); var b = E(doc, "e"); b.AppendChild(atom); script.AppendChild(b);
            if (n.Sub != null) script.AppendChild(Container(doc, "sub", n.Sub, pt));
            if (n.Sup != null) script.AppendChild(Container(doc, "sup", n.Sup, pt));
            return script;
        }
        static XmlElement Atom(XmlDocument doc, TexNode n, double pt) {
            XmlElement a;
            if (n.Kind == "text") {
                a = E(doc, "r"); var rp = E(doc, "rPr");
                if (n.Style == "script") rp.AppendChild(E(doc, "scr", "script"));
                else if (n.Style == "plain" || n.Style == "bold" || n.Style == "italic") rp.AppendChild(E(doc, "sty", n.Style == "plain" ? "p" : n.Style == "bold" ? "b" : "i"));
                a.AppendChild(rp);
                var wrp = doc.CreateElement("w", "rPr", W);
                var fonts = doc.CreateElement("w", "rFonts", W); fonts.SetAttribute("ascii", W, "Cambria Math"); fonts.SetAttribute("hAnsi", W, "Cambria Math"); fonts.SetAttribute("eastAsia", W, "宋体"); wrp.AppendChild(fonts);
                var size = doc.CreateElement("w", "sz", W); size.SetAttribute("val", W, Math.Round(pt * 2).ToString(System.Globalization.CultureInfo.InvariantCulture)); wrp.AppendChild(size); a.AppendChild(wrp);
                var t = E(doc, "t"); t.InnerText = n.Value; t.SetAttribute("space", "http://www.w3.org/XML/1998/namespace", "preserve"); a.AppendChild(t);
            } else if (n.Kind == "group") { a = E(doc, "box"); a.AppendChild(Container(doc, "e", n.Children, pt)); }
            else if (n.Kind == "fraction") { a = E(doc, "f"); a.AppendChild(Container(doc, "num", n.Children, pt)); a.AppendChild(Container(doc, "den", n.Sub, pt)); }
            else if (n.Kind == "radical") { a = E(doc, "rad"); var p = E(doc, "radPr"); p.AppendChild(E(doc, "degHide", "1")); a.AppendChild(p); a.AppendChild(E(doc,"deg")); a.AppendChild(Container(doc,"e",n.Children,pt)); }
            else if (n.Kind == "accent") { a = E(doc, "acc"); var p = E(doc,"accPr"); p.AppendChild(E(doc,"chr",n.Value)); a.AppendChild(p); a.AppendChild(Container(doc,"e",n.Children,pt)); }
            else if (n.Kind == "delimiter") { a = E(doc,"d"); var p = E(doc,"dPr"); string[] delimiters = n.Value.Split('\u0001'); p.AppendChild(E(doc,"begChr",delimiters[0])); p.AppendChild(E(doc,"endChr",delimiters[1])); a.AppendChild(p); a.AppendChild(Container(doc,"e",n.Children,pt)); }
            else if (n.Kind == "cases") {
                a = E(doc,"d"); var dp=E(doc,"dPr"); dp.AppendChild(E(doc,"begChr","{")); dp.AppendChild(E(doc,"endChr","")); a.AppendChild(dp);
                var expr=E(doc,"e"); var matrix=E(doc,"m");
                foreach(var row in n.Rows) { var mr=E(doc,"mr"); foreach(var cell in row) mr.AppendChild(Container(doc,"e",cell,pt)); matrix.AppendChild(mr); }
                expr.AppendChild(matrix); a.AppendChild(expr);
            } else throw new FormatException("Unsupported OMML node " + n.Kind);
            return Script(doc,n,a,pt);
        }
        static void Append(XmlDocument doc, XmlElement parent, List<TexNode> nodes, double pt) {
            if (nodes == null) return;
            for(int i=0;i<nodes.Count;i++) {
                TexNode n=nodes[i];
                if(n.Kind == "sum") {
                    var sum=E(doc,"nary"); var p=E(doc,"naryPr"); p.AppendChild(E(doc,"chr","∑")); p.AppendChild(E(doc,"limLoc","subSup"));
                    if(n.Sub == null) p.AppendChild(E(doc,"subHide","1")); if(n.Sup == null) p.AppendChild(E(doc,"supHide","1")); sum.AppendChild(p);
                    sum.AppendChild(Container(doc,"sub",n.Sub,pt)); sum.AppendChild(Container(doc,"sup",n.Sup,pt));
                    // Summations in the supported manuscript bind the remaining product in this group.
                    sum.AppendChild(Container(doc,"e",nodes.GetRange(i+1,nodes.Count-i-1),pt)); parent.AppendChild(sum); return;
                }
                parent.AppendChild(Atom(doc,n,pt));
            }
        }
        public static XmlElement Convert(XmlDocument doc, string tex, double pt) {
            var result=E(doc,"oMath"); Append(doc,result,new TexParser(tex).Parse(),pt); return result;
        }
        static string ReadableSequence(List<TexNode> nodes) {
            if(nodes == null) return "";
            var text=new StringBuilder();
            foreach(var n in nodes) {
                string value;
                if(n.Kind == "text") value=n.Value;
                else if(n.Kind == "group") value=ReadableSequence(n.Children);
                else if(n.Kind == "fraction") value="("+ReadableSequence(n.Children)+")/("+ReadableSequence(n.Sub)+")";
                else if(n.Kind == "radical") value="√("+ReadableSequence(n.Children)+")";
                else if(n.Kind == "accent") value=(n.Value == "̂" ? "hat" : n.Value == "̅" ? "bar" : "tilde")+"("+ReadableSequence(n.Children)+")";
                else if(n.Kind == "sum") value="Σ";
                else if(n.Kind == "delimiter") {var d=n.Value.Split('\u0001');value=d[0]+ReadableSequence(n.Children)+d[1];}
                else if(n.Kind == "cases") {
                    var rows=new List<string>();foreach(var row in n.Rows){var cells=new List<string>();foreach(var cell in row)cells.Add(ReadableSequence(cell));rows.Add(String.Join(" ",cells));}
                    value="cases{"+String.Join("; ",rows)+"}";
                }else throw new FormatException("Unsupported readable node "+n.Kind);
                if(n.Kind != "fraction" && n.Sub != null)value+="_{"+ReadableSequence(n.Sub)+"}";
                if(n.Sup != null)value+="^{"+ReadableSequence(n.Sup)+"}";
                text.Append(value);
            }
            return text.ToString();
        }
        public static string Readable(string tex) { return ReadableSequence(new TexParser(tex).Parse()); }
    }
}
