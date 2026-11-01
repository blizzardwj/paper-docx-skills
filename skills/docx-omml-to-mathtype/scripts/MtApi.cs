// Minimal P/Invoke wrapper around MathType 7's 64-bit API (MT6.dll).
// Every quirk handled here was found the hard way; see references/pitfalls.md before changing anything.
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class MtApi {
    const string Dll = @"C:\Program Files (x86)\MathType\System\64\MT6.dll";

    [DllImport(Dll)] public static extern int MTAPIConnect(short options, short timeout);
    [DllImport(Dll)] public static extern int MTAPIDisconnect();
    [DllImport(Dll)] public static extern int MTXFormReset();
    [DllImport(Dll)] public static extern int MTEquationOnClipboard();   // 8 = OLE2 equation
    [DllImport(Dll, CharSet = CharSet.Ansi)] static extern int MTGetPrefsFromFile(string prefFile, StringBuilder prefs, short prefsLen);
    [DllImport(Dll)] static extern int MTXFormSetPrefs(short prefType, IntPtr prefStr);
    [DllImport(Dll, EntryPoint = "MTXFormEqn")]
    static extern int MTXFormEqnRaw(short src, short srcFmt, IntPtr srcData, int srcLen,
                                    short dst, short dstFmt, IntPtr dstData, int dstLen,
                                    IntPtr dstPath, IntPtr dims);

    public const short CLIPBOARD = -2, FILE = -4;
    public const short PICT = 6, TEXT = 7, GIF = 9;
    const short PREF_USER = 3;

    public static string LoadPrefs(string eqpPath) {
        int n = MTGetPrefsFromFile(eqpPath, null, 0);
        if (n <= 0) throw new Exception("MTGetPrefsFromFile size failed: " + n);
        var sb = new StringBuilder(n + 16);
        int rc = MTGetPrefsFromFile(eqpPath, sb, (short)(n + 16));
        if (rc < 0) throw new Exception("MTGetPrefsFromFile failed: " + rc);
        return sb.ToString();
    }

    // MathType keeps the prefs pointer until the next MTXFormEqn, so the string must outlive the call.
    static IntPtr prefsMem = IntPtr.Zero;
    static string prefsMemFor = null;
    static int SetPrefs(string prefs) {
        if (prefs == null) return 0;
        if (prefs != prefsMemFor) {
            if (prefsMem != IntPtr.Zero) Marshal.FreeHGlobal(prefsMem);
            prefsMem = Marshal.StringToHGlobalAnsi(prefs);
            prefsMemFor = prefs;
        }
        return MTXFormSetPrefs(PREF_USER, prefsMem);
    }

    // The native side writes more than the VBA-declared structs hold, so every buffer is a generous
    // unmanaged block (small managed buffers caused random access violations).
    static int Xform(short src, short srcFmt, short dst, short dstFmt, string path) {
        const int BIG = 1 << 16;
        IntPtr srcBuf = Marshal.AllocHGlobal(BIG), dstBuf = Marshal.AllocHGlobal(BIG), dimsBuf = Marshal.AllocHGlobal(4096);
        IntPtr pathBuf = Marshal.StringToHGlobalAnsi((path ?? " ") + new string('\0', 8));
        try {
            foreach (var p in new[] { srcBuf, dstBuf }) { for (int i = 0; i < 64; i++) Marshal.WriteByte(p, i, 0); Marshal.WriteByte(p, 0, 32); }
            for (int i = 0; i < 4096; i++) Marshal.WriteByte(dimsBuf, i, 0);
            return MTXFormEqnRaw(src, srcFmt, srcBuf, 1, dst, dstFmt, dstBuf, 1, pathBuf, dimsBuf);
        } finally {
            Marshal.FreeHGlobal(srcBuf); Marshal.FreeHGlobal(dstBuf); Marshal.FreeHGlobal(dimsBuf); Marshal.FreeHGlobal(pathBuf);
        }
    }

    // MathML text already on the clipboard -> MathType equation (dst CLIPBOARD+PICT = OLE equation object).
    public static int ConvertClipboardMathML(string prefs, short dst, short dstFmt, string path) {
        MTXFormReset();
        int p = SetPrefs(prefs); if (p < 0) return p;
        return Xform(CLIPBOARD, TEXT, dst, dstFmt, path);
    }

    // Render the MathType equation currently on the clipboard to a file, leaving the clipboard as is.
    public static int RenderClipboardEquation(string prefs, short dstFmt, string path) {
        MTXFormReset();
        int p = SetPrefs(prefs); if (p < 0) return p;
        return Xform(CLIPBOARD, PICT, FILE, dstFmt, path);
    }
}
