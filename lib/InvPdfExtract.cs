// ============================================================================
//  InvPdfExtract.cs  --  PDF text extraction with coordinates.
//
//  NO INSTALL, NO DOWNLOAD, NO DLL.
//    This is C# SOURCE, compiled at run time by Add-Type using the compiler
//    that ships inside the .NET Framework on every Windows machine -- the same
//    trick NeXlsx.ps1 already uses in the address repo. Nothing is fetched and
//    nothing is installed. Ever.
//
//  WHY IT EXISTS
//    Invoice layouts are all different, so we can never parse "the third line".
//    We need every word AND where it sits, so a field can be found by what it
//    is next to ("Invoice Date" -> the value to its right).
//
//  WHY C# AND NOT POWERSHELL
//    A PDF is binary and a page's content stream is tens of thousands of
//    tokens. Per-token work in PowerShell would put a 500-invoice run into the
//    minutes; here it stays in the tens of milliseconds per page.
//
//  PIPELINE
//    bytes -> objects -> decoded streams -> page tree -> content operators
//          -> positioned glyphs -> words with X/Y/W/H in top-left points
//
//  LANGUAGE LEVEL -- C# 5 ONLY
//    Windows PowerShell 5.1's Add-Type uses the .NET Framework CodeDom
//    compiler, which predates C# 6. So:
//      NO string interpolation  NO ?. or ??=   NO nameof   NO tuples
//      NO expression-bodied members            NO out var
//      NO auto-property initialisers           NO using static
//    tests/Test-LangLevel.ps1 compiles this file at LangVersion 5 so a slip is
//    caught here rather than on your machine.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace InvParse
{
    // ------------------------------------------------------------- the model

    public abstract class PdfObj { }

    public sealed class PdfNull : PdfObj
    {
        public static readonly PdfNull Instance = new PdfNull();
    }

    public sealed class PdfBool : PdfObj
    {
        public bool Value;
        public PdfBool(bool v) { Value = v; }
    }

    public sealed class PdfNum : PdfObj
    {
        public double Value;
        public PdfNum(double v) { Value = v; }
    }

    /// <summary>A PDF string, kept as raw bytes. Text drawn on the page is
    /// decoded through the font's own encoding later, never as ASCII.</summary>
    public sealed class PdfStr : PdfObj
    {
        public byte[] Bytes;
        public PdfStr(byte[] b) { Bytes = b; }
    }

    public sealed class PdfName : PdfObj
    {
        public string Value;
        public PdfName(string v) { Value = v; }
    }

    public sealed class PdfArray : PdfObj
    {
        public List<PdfObj> Items;
        public PdfArray() { Items = new List<PdfObj>(); }
    }

    public sealed class PdfDict : PdfObj
    {
        public Dictionary<string, PdfObj> Map;
        public PdfDict() { Map = new Dictionary<string, PdfObj>(StringComparer.Ordinal); }

        public PdfObj Get(string key)
        {
            PdfObj v;
            if (Map.TryGetValue(key, out v)) return v;
            return null;
        }
    }

    public sealed class PdfStream : PdfObj
    {
        public PdfDict Dict;
        public byte[] Raw;
        public byte[] Decoded;
        public bool Tried;
        public PdfStream() { Dict = new PdfDict(); }
    }

    public sealed class PdfRef : PdfObj
    {
        public int Num;
        public int Gen;
        public PdfRef(int n, int g) { Num = n; Gen = g; }
    }

    // -------------------------------------------------------------- the lexer

    /// <summary>
    /// Tokenising parser over a byte buffer. Serves both file objects and page
    /// content streams -- the syntax is identical, content streams simply have
    /// operators where the file body has obj/endobj.
    /// </summary>
    public sealed class PdfLexer
    {
        public byte[] Buf;
        public int Pos;
        public string LastKeyword;

        public PdfLexer(byte[] buf, int pos) { Buf = buf; Pos = pos; }

        public static bool IsWhite(byte b)
        {
            return b == 0 || b == 9 || b == 10 || b == 12 || b == 13 || b == 32;
        }

        public static bool IsDelim(byte b)
        {
            return b == (byte)'(' || b == (byte)')' || b == (byte)'<' || b == (byte)'>'
                || b == (byte)'[' || b == (byte)']' || b == (byte)'{' || b == (byte)'}'
                || b == (byte)'/' || b == (byte)'%';
        }

        public static bool IsRegular(byte b) { return !IsWhite(b) && !IsDelim(b); }

        public void SkipWhite()
        {
            while (Pos < Buf.Length)
            {
                byte b = Buf[Pos];
                if (IsWhite(b)) { Pos++; continue; }
                if (b == (byte)'%')
                {
                    while (Pos < Buf.Length && Buf[Pos] != 10 && Buf[Pos] != 13) Pos++;
                    continue;
                }
                break;
            }
        }

        public string ReadKeyword()
        {
            SkipWhite();
            int start = Pos;
            while (Pos < Buf.Length && IsRegular(Buf[Pos])) Pos++;
            if (Pos == start && Pos < Buf.Length) Pos++;
            return Encoding.ASCII.GetString(Buf, start, Pos - start);
        }

        private static int HexVal(byte b)
        {
            if (b >= (byte)'0' && b <= (byte)'9') return b - (byte)'0';
            if (b >= (byte)'a' && b <= (byte)'f') return b - (byte)'a' + 10;
            if (b >= (byte)'A' && b <= (byte)'F') return b - (byte)'A' + 10;
            return -1;
        }

        private string ReadNameToken()
        {
            StringBuilder sb = new StringBuilder();
            while (Pos < Buf.Length && IsRegular(Buf[Pos]))
            {
                byte b = Buf[Pos++];
                if (b == (byte)'#' && Pos + 1 < Buf.Length)
                {
                    int hi = HexVal(Buf[Pos]);
                    int lo = HexVal(Buf[Pos + 1]);
                    if (hi >= 0 && lo >= 0) { sb.Append((char)(hi * 16 + lo)); Pos += 2; continue; }
                }
                sb.Append((char)b);
            }
            return sb.ToString();
        }

        private PdfObj ReadLiteralString()
        {
            List<byte> o = new List<byte>();
            int depth = 1;
            while (Pos < Buf.Length)
            {
                byte b = Buf[Pos++];
                if (b == (byte)'\\')
                {
                    if (Pos >= Buf.Length) break;
                    byte e = Buf[Pos++];
                    switch ((char)e)
                    {
                        case 'n': o.Add(10); break;
                        case 'r': o.Add(13); break;
                        case 't': o.Add(9); break;
                        case 'b': o.Add(8); break;
                        case 'f': o.Add(12); break;
                        case '(': o.Add((byte)'('); break;
                        case ')': o.Add((byte)')'); break;
                        case '\\': o.Add((byte)'\\'); break;
                        case '\r':
                            if (Pos < Buf.Length && Buf[Pos] == 10) Pos++;
                            break;
                        case '\n': break;
                        default:
                            if (e >= (byte)'0' && e <= (byte)'7')
                            {
                                int v = e - (byte)'0';
                                for (int i = 0; i < 2 && Pos < Buf.Length; i++)
                                {
                                    byte d = Buf[Pos];
                                    if (d < (byte)'0' || d > (byte)'7') break;
                                    v = v * 8 + (d - (byte)'0');
                                    Pos++;
                                }
                                o.Add((byte)(v & 0xFF));
                            }
                            else o.Add(e);
                            break;
                    }
                    continue;
                }
                if (b == (byte)'(') { depth++; o.Add(b); continue; }
                if (b == (byte)')')
                {
                    depth--;
                    if (depth == 0) break;
                    o.Add(b);
                    continue;
                }
                o.Add(b);
            }
            return new PdfStr(o.ToArray());
        }

        private PdfObj ReadHexString()
        {
            List<byte> o = new List<byte>();
            int hi = -1;
            while (Pos < Buf.Length)
            {
                byte b = Buf[Pos++];
                if (b == (byte)'>') break;
                int v = HexVal(b);
                if (v < 0) continue;
                if (hi < 0) hi = v;
                else { o.Add((byte)(hi * 16 + v)); hi = -1; }
            }
            if (hi >= 0) o.Add((byte)(hi * 16));
            return new PdfStr(o.ToArray());
        }

        /// <summary>Parses one object. Returns null on an operator token, which
        /// is left in LastKeyword, or at end of buffer.</summary>
        public PdfObj ParseObject()
        {
            LastKeyword = null;
            SkipWhite();
            if (Pos >= Buf.Length) return null;

            byte b = Buf[Pos];

            if (b == (byte)'/') { Pos++; return new PdfName(ReadNameToken()); }
            if (b == (byte)'(') { Pos++; return ReadLiteralString(); }

            if (b == (byte)'<')
            {
                if (Pos + 1 < Buf.Length && Buf[Pos + 1] == (byte)'<') { Pos += 2; return ParseDictBody(); }
                Pos++;
                return ReadHexString();
            }

            if (b == (byte)'[')
            {
                Pos++;
                PdfArray arr = new PdfArray();
                while (true)
                {
                    SkipWhite();
                    if (Pos >= Buf.Length) break;
                    if (Buf[Pos] == (byte)']') { Pos++; break; }
                    int before = Pos;
                    PdfObj it = ParseObject();
                    if (it == null)
                    {
                        if (Pos == before) Pos++;          // never spin
                        continue;
                    }
                    arr.Items.Add(it);
                }
                return arr;
            }

            if (b == (byte)']' || b == (byte)'>' || b == (byte)')' || b == (byte)'}' || b == (byte)'{')
            {
                Pos++;
                return null;
            }

            if ((b >= (byte)'0' && b <= (byte)'9') || b == (byte)'+' || b == (byte)'-' || b == (byte)'.')
            {
                string tok = ReadKeyword();
                double d;
                if (!TryNum(tok, out d)) return new PdfNum(0);

                if (IsInt(tok))                            // maybe "<n> <g> R"
                {
                    int save = Pos;
                    string t2 = ReadKeyword();
                    if (IsInt(t2))
                    {
                        string t3 = ReadKeyword();
                        if (t3 == "R")
                            return new PdfRef((int)d, int.Parse(t2, CultureInfo.InvariantCulture));
                    }
                    Pos = save;
                }
                return new PdfNum(d);
            }

            string kw = ReadKeyword();
            if (kw == "true") return new PdfBool(true);
            if (kw == "false") return new PdfBool(false);
            if (kw == "null") return PdfNull.Instance;
            if (kw.Length == 0) { Pos++; return null; }

            LastKeyword = kw;
            return null;
        }

        public PdfDict ParseDictBody()
        {
            PdfDict d = new PdfDict();
            while (true)
            {
                SkipWhite();
                if (Pos >= Buf.Length) break;
                if (Buf[Pos] == (byte)'>')
                {
                    Pos++;
                    if (Pos < Buf.Length && Buf[Pos] == (byte)'>') Pos++;
                    break;
                }
                if (Buf[Pos] != (byte)'/')
                {
                    int before = Pos;
                    PdfObj junk = ParseObject();
                    if (junk == null && Pos == before) Pos++;
                    continue;
                }
                Pos++;
                string key = ReadNameToken();
                int b2 = Pos;
                PdfObj val = ParseObject();
                if (val == null)
                {
                    if (Pos == b2) Pos++;
                    continue;
                }
                d.Map[key] = val;
            }
            return d;
        }

        public static bool IsInt(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            for (int i = 0; i < s.Length; i++) if (s[i] < '0' || s[i] > '9') return false;
            return true;
        }

        /// <summary>PDF tolerates numbers .NET rejects: "--5", "4.", ".5". Parse
        /// leniently; a malformed number must never sink a whole invoice.</summary>
        public static bool TryNum(string s, out double d)
        {
            d = 0;
            if (string.IsNullOrEmpty(s)) return false;
            StringBuilder sb = new StringBuilder();
            bool neg = false, dot = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '-') { if (sb.Length == 0) neg = !neg; continue; }
                if (c == '+') continue;
                if (c == '.') { if (dot) break; dot = true; sb.Append('.'); continue; }
                if (c < '0' || c > '9') break;
                sb.Append(c);
            }
            string t = sb.ToString();
            if (t.Length > 0 && t[t.Length - 1] == '.') t = t.Substring(0, t.Length - 1);
            if (t.Length == 0 || t == ".") return false;
            if (!double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return false;
            if (neg) d = -d;
            return true;
        }
    }

    // ------------------------------------------------------------- filters

    public static class PdfFilters
    {
        /// <summary>Inflate. PDFs carry zlib-wrapped deflate; some writers emit
        /// raw deflate, and a few prepend junk. Try the likely starts in turn
        /// and keep whatever produces the most output.</summary>
        public static byte[] Flate(byte[] data)
        {
            if (data == null || data.Length == 0) return new byte[0];

            int[] starts;
            if (data.Length > 2 && data[0] == 0x78) starts = new int[] { 2, 0, 1 };
            else starts = new int[] { 0, 2, 1 };

            byte[] best = null;
            for (int si = 0; si < starts.Length; si++)
            {
                int s = starts[si];
                if (s >= data.Length) continue;
                byte[] got = TryInflate(data, s);
                if (got != null && (best == null || got.Length > best.Length)) best = got;
                if (best != null && best.Length > 0 && si == 0) break;   // first guess worked
            }
            if (best == null) return new byte[0];
            return best;
        }

        private static byte[] TryInflate(byte[] data, int offset)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream(data, offset, data.Length - offset))
                using (DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress))
                using (MemoryStream outp = new MemoryStream())
                {
                    byte[] buf = new byte[16384];
                    int n;
                    // A truncated stream still yields everything before the break,
                    // which is usually the whole page. Keep what we got.
                    try
                    {
                        while ((n = ds.Read(buf, 0, buf.Length)) > 0) outp.Write(buf, 0, n);
                    }
                    catch (Exception) { }
                    return outp.ToArray();
                }
            }
            catch (Exception) { return null; }
        }

        public static byte[] AsciiHex(byte[] data)
        {
            List<byte> o = new List<byte>();
            int hi = -1;
            for (int i = 0; i < data.Length; i++)
            {
                byte b = data[i];
                if (b == (byte)'>') break;
                int v = -1;
                if (b >= (byte)'0' && b <= (byte)'9') v = b - (byte)'0';
                else if (b >= (byte)'a' && b <= (byte)'f') v = b - (byte)'a' + 10;
                else if (b >= (byte)'A' && b <= (byte)'F') v = b - (byte)'A' + 10;
                if (v < 0) continue;
                if (hi < 0) hi = v; else { o.Add((byte)(hi * 16 + v)); hi = -1; }
            }
            if (hi >= 0) o.Add((byte)(hi * 16));
            return o.ToArray();
        }

        public static byte[] Ascii85(byte[] data)
        {
            List<byte> o = new List<byte>();
            uint tuple = 0;
            int count = 0;
            int i = 0;
            if (data.Length > 1 && data[0] == (byte)'<' && data[1] == (byte)'~') i = 2;
            for (; i < data.Length; i++)
            {
                byte b = data[i];
                if (b == (byte)'~') break;
                if (PdfLexer.IsWhite(b)) continue;
                if (b == (byte)'z' && count == 0) { o.Add(0); o.Add(0); o.Add(0); o.Add(0); continue; }
                if (b < 33 || b > 117) continue;
                tuple = tuple * 85 + (uint)(b - 33);
                count++;
                if (count == 5)
                {
                    o.Add((byte)(tuple >> 24)); o.Add((byte)(tuple >> 16));
                    o.Add((byte)(tuple >> 8)); o.Add((byte)tuple);
                    tuple = 0; count = 0;
                }
            }
            if (count > 0)
            {
                for (int k = count; k < 5; k++) tuple = tuple * 85 + 84;
                for (int k = 0; k < count - 1; k++) o.Add((byte)(tuple >> (24 - 8 * k)));
            }
            return o.ToArray();
        }

        public static byte[] RunLength(byte[] data)
        {
            List<byte> o = new List<byte>();
            int i = 0;
            while (i < data.Length)
            {
                int len = data[i++];
                if (len == 128) break;
                if (len < 128)
                {
                    for (int k = 0; k <= len && i < data.Length; k++) o.Add(data[i++]);
                }
                else
                {
                    if (i >= data.Length) break;
                    byte b = data[i++];
                    for (int k = 0; k < 257 - len; k++) o.Add(b);
                }
            }
            return o.ToArray();
        }

        public static byte[] Lzw(byte[] data, int early)
        {
            List<byte> o = new List<byte>();
            byte[][] table = new byte[4096][];
            for (int i = 0; i < 256; i++) table[i] = new byte[] { (byte)i };
            int next = 258;
            int codeLen = 9;
            byte[] prev = null;

            int bitPos = 0;
            int totalBits = data.Length * 8;
            while (bitPos + codeLen <= totalBits)
            {
                int code = 0;
                for (int k = 0; k < codeLen; k++)
                {
                    int bp = bitPos + k;
                    int bit = (data[bp >> 3] >> (7 - (bp & 7))) & 1;
                    code = (code << 1) | bit;
                }
                bitPos += codeLen;

                if (code == 256)
                {
                    next = 258; codeLen = 9; prev = null;
                    for (int i = 258; i < 4096; i++) table[i] = null;
                    continue;
                }
                if (code == 257) break;

                byte[] entry;
                if (code < next && table[code] != null) entry = table[code];
                else if (prev != null)
                {
                    entry = new byte[prev.Length + 1];
                    Array.Copy(prev, entry, prev.Length);
                    entry[prev.Length] = prev[0];
                }
                else break;

                o.AddRange(entry);

                if (prev != null && next < 4096)
                {
                    byte[] ne = new byte[prev.Length + 1];
                    Array.Copy(prev, ne, prev.Length);
                    ne[prev.Length] = entry[0];
                    table[next++] = ne;
                }
                prev = entry;

                int limit = next + (early != 0 ? 1 : 0);
                if (limit >= 512 && codeLen == 9) codeLen = 10;
                else if (limit >= 1024 && codeLen == 10) codeLen = 11;
                else if (limit >= 2048 && codeLen == 11) codeLen = 12;
            }
            return o.ToArray();
        }

        /// <summary>Undo PNG/TIFF predictors (used by xref and object streams).</summary>
        public static byte[] Predictor(byte[] data, int pred, int colors, int bpc, int columns)
        {
            if (pred < 2) return data;
            int bpp = Math.Max(1, (colors * bpc + 7) / 8);
            int rowLen = (columns * colors * bpc + 7) / 8;

            if (pred == 2)
            {
                if (bpc != 8) return data;
                for (int r = 0; r + rowLen <= data.Length; r += rowLen)
                    for (int i = bpp; i < rowLen; i++)
                        data[r + i] = (byte)((data[r + i] + data[r + i - bpp]) & 0xFF);
                return data;
            }

            // PNG predictors: each row is prefixed with a filter-type byte
            int rows = data.Length / (rowLen + 1);
            byte[] outp = new byte[rows * rowLen];
            byte[] prior = new byte[rowLen];
            int src = 0, dst = 0;
            for (int r = 0; r < rows; r++)
            {
                int ft = data[src++];
                byte[] cur = new byte[rowLen];
                Array.Copy(data, src, cur, 0, rowLen);
                src += rowLen;

                for (int i = 0; i < rowLen; i++)
                {
                    int a = (i >= bpp) ? cur[i - bpp] : 0;
                    int b = prior[i];
                    int c = (i >= bpp) ? prior[i - bpp] : 0;
                    int x = cur[i];
                    int v;
                    switch (ft)
                    {
                        case 0: v = x; break;
                        case 1: v = x + a; break;
                        case 2: v = x + b; break;
                        case 3: v = x + ((a + b) >> 1); break;
                        case 4:
                            int p = a + b - c;
                            int pa = Math.Abs(p - a), pb = Math.Abs(p - b), pc = Math.Abs(p - c);
                            int pr = (pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c);
                            v = x + pr;
                            break;
                        default: v = x; break;
                    }
                    cur[i] = (byte)(v & 0xFF);
                }
                Array.Copy(cur, 0, outp, dst, rowLen);
                dst += rowLen;
                prior = cur;
            }
            return outp;
        }
    }

    // ------------------------------------------------------------ the document

    /// <summary>
    /// The parsed file. Objects are found by SCANNING for "N G obj" rather than
    /// by following the xref table: real-world invoices come out of every
    /// printer driver and mail merge under the sun, and a broken or lying xref
    /// is common. Scanning does not care, and incremental updates resolve
    /// naturally because a later definition of an object wins.
    /// </summary>
    public sealed class PdfDoc
    {
        public byte[] Buf;
        public Dictionary<int, PdfObj> Objects;
        public List<PdfDict> Pages;
        public bool Encrypted;

        public PdfDoc()
        {
            Objects = new Dictionary<int, PdfObj>();
            Pages = new List<PdfDict>();
        }

        public static PdfDoc Load(string path)
        {
            PdfDoc doc = new PdfDoc();
            doc.Buf = File.ReadAllBytes(path);
            doc.ScanObjects();
            doc.ExpandObjectStreams();
            doc.DetectEncryption();
            doc.BuildPageList();
            return doc;
        }

        public PdfObj Resolve(PdfObj o)
        {
            int guard = 0;
            while (o is PdfRef && guard++ < 64)
            {
                PdfObj v;
                if (!Objects.TryGetValue(((PdfRef)o).Num, out v)) return null;
                o = v;
            }
            return o;
        }

        public PdfDict ResolveDict(PdfObj o)
        {
            PdfObj r = Resolve(o);
            PdfStream st = r as PdfStream;
            if (st != null) return st.Dict;
            return r as PdfDict;
        }

        public double GetNum(PdfDict d, string key, double dflt)
        {
            if (d == null) return dflt;
            PdfNum n = Resolve(d.Get(key)) as PdfNum;
            if (n == null) return dflt;
            return n.Value;
        }

        public string GetName(PdfDict d, string key)
        {
            if (d == null) return null;
            PdfName n = Resolve(d.Get(key)) as PdfName;
            if (n == null) return null;
            return n.Value;
        }

        // ---- object scan ----

        private void ScanObjects()
        {
            byte[] b = Buf;
            for (int i = 0; i + 2 < b.Length; i++)
            {
                if (b[i] != (byte)'o' || b[i + 1] != (byte)'b' || b[i + 2] != (byte)'j') continue;
                if (i + 3 < b.Length && PdfLexer.IsRegular(b[i + 3])) continue;
                if (i == 0 || !PdfLexer.IsWhite(b[i - 1])) continue;

                // walk back over "<num> <gen> "
                int p = i - 1;
                while (p >= 0 && PdfLexer.IsWhite(b[p])) p--;
                int genEnd = p + 1;
                while (p >= 0 && b[p] >= (byte)'0' && b[p] <= (byte)'9') p--;
                int genStart = p + 1;
                if (genStart >= genEnd) continue;
                while (p >= 0 && PdfLexer.IsWhite(b[p])) p--;
                int numEnd = p + 1;
                while (p >= 0 && b[p] >= (byte)'0' && b[p] <= (byte)'9') p--;
                int numStart = p + 1;
                if (numStart >= numEnd) continue;
                if (numEnd == genEnd) continue;
                if (p >= 0 && PdfLexer.IsRegular(b[p])) continue;

                int num;
                if (!int.TryParse(Encoding.ASCII.GetString(b, numStart, numEnd - numStart),
                                  out num)) continue;

                PdfLexer lex = new PdfLexer(b, i + 3);
                PdfObj val = lex.ParseObject();
                if (val == null && lex.LastKeyword == null) continue;

                // a stream body may follow the dictionary
                PdfDict dict = val as PdfDict;
                if (dict != null)
                {
                    int save = lex.Pos;
                    string kw = lex.ReadKeyword();
                    if (kw == "stream")
                    {
                        int dp = lex.Pos;
                        if (dp < b.Length && b[dp] == 13) dp++;
                        if (dp < b.Length && b[dp] == 10) dp++;
                        int len = StreamLength(dict, dp);
                        PdfStream st = new PdfStream();
                        st.Dict = dict;
                        st.Raw = new byte[len];
                        if (len > 0) Array.Copy(b, dp, st.Raw, 0, len);
                        Objects[num] = st;
                        i = dp + len;
                        continue;
                    }
                    lex.Pos = save;
                }

                if (val != null) Objects[num] = val;
            }
        }

        /// <summary>/Length is often an indirect reference we have not read yet,
        /// and is sometimes simply wrong. Trust it only when the bytes it points
        /// at are actually followed by "endstream".</summary>
        private int StreamLength(PdfDict dict, int dataStart)
        {
            PdfObj lenObj = dict.Get("Length");
            PdfNum ln = lenObj as PdfNum;
            if (ln != null)
            {
                int len = (int)ln.Value;
                if (len >= 0 && dataStart + len <= Buf.Length && EndstreamNear(dataStart + len))
                    return len;
            }
            int idx = IndexOf(Buf, Encoding.ASCII.GetBytes("endstream"), dataStart);
            if (idx < 0) return Math.Max(0, Buf.Length - dataStart);
            int end = idx;
            if (end > dataStart && Buf[end - 1] == 10) end--;
            if (end > dataStart && Buf[end - 1] == 13) end--;
            return end - dataStart;
        }

        private bool EndstreamNear(int pos)
        {
            for (int k = pos; k < Math.Min(Buf.Length - 8, pos + 4); k++)
            {
                if (Buf[k] == (byte)'e' && Buf[k + 1] == (byte)'n' && Buf[k + 2] == (byte)'d'
                    && Buf[k + 3] == (byte)'s' && Buf[k + 4] == (byte)'t') return true;
            }
            return false;
        }

        public static int IndexOf(byte[] hay, byte[] needle, int from)
        {
            int last = hay.Length - needle.Length;
            for (int i = Math.Max(0, from); i <= last; i++)
            {
                int k = 0;
                while (k < needle.Length && hay[i + k] == needle[k]) k++;
                if (k == needle.Length) return i;
            }
            return -1;
        }

        // ---- stream data ----

        public byte[] GetStreamData(PdfStream st)
        {
            if (st == null) return new byte[0];
            if (st.Tried) return st.Decoded == null ? new byte[0] : st.Decoded;
            st.Tried = true;

            byte[] data = st.Raw;
            PdfObj f = Resolve(st.Dict.Get("Filter"));
            PdfObj parms = Resolve(st.Dict.Get("DecodeParms"));
            if (parms == null) parms = Resolve(st.Dict.Get("DP"));

            List<string> filters = new List<string>();
            PdfName fn = f as PdfName;
            if (fn != null) filters.Add(fn.Value);
            PdfArray fa = f as PdfArray;
            if (fa != null)
                for (int i = 0; i < fa.Items.Count; i++)
                {
                    PdfName x = Resolve(fa.Items[i]) as PdfName;
                    if (x != null) filters.Add(x.Value);
                }

            for (int i = 0; i < filters.Count; i++)
            {
                PdfDict pd = null;
                PdfArray pa = parms as PdfArray;
                if (pa != null && i < pa.Items.Count) pd = ResolveDict(pa.Items[i]);
                else if (i == 0) pd = parms as PdfDict;

                switch (filters[i])
                {
                    case "FlateDecode":
                    case "Fl":
                        data = PdfFilters.Flate(data);
                        data = ApplyPredictor(data, pd);
                        break;
                    case "LZWDecode":
                    case "LZW":
                        data = PdfFilters.Lzw(data, (int)GetNum(pd, "EarlyChange", 1));
                        data = ApplyPredictor(data, pd);
                        break;
                    case "ASCIIHexDecode":
                    case "AHx":
                        data = PdfFilters.AsciiHex(data);
                        break;
                    case "ASCII85Decode":
                    case "A85":
                        data = PdfFilters.Ascii85(data);
                        break;
                    case "RunLengthDecode":
                    case "RL":
                        data = PdfFilters.RunLength(data);
                        break;
                    default:
                        // DCTDecode / JPXDecode / CCITTFaxDecode / JBIG2Decode are
                        // images. There is no text in them; hand back nothing.
                        st.Decoded = new byte[0];
                        return st.Decoded;
                }
            }
            st.Decoded = data;
            return data;
        }

        private byte[] ApplyPredictor(byte[] data, PdfDict pd)
        {
            if (pd == null) return data;
            int pred = (int)GetNum(pd, "Predictor", 1);
            if (pred < 2) return data;
            return PdfFilters.Predictor(data, pred,
                (int)GetNum(pd, "Colors", 1),
                (int)GetNum(pd, "BitsPerComponent", 8),
                (int)GetNum(pd, "Columns", 1));
        }

        // ---- object streams (PDF 1.5+ packs objects inside a stream) ----

        private void ExpandObjectStreams()
        {
            List<PdfStream> objStms = new List<PdfStream>();
            foreach (KeyValuePair<int, PdfObj> kv in Objects)
            {
                PdfStream st = kv.Value as PdfStream;
                if (st == null) continue;
                PdfName t = st.Dict.Get("Type") as PdfName;
                if (t != null && t.Value == "ObjStm") objStms.Add(st);
            }

            for (int s = 0; s < objStms.Count; s++)
            {
                PdfStream st = objStms[s];
                byte[] data = GetStreamData(st);
                if (data.Length == 0) continue;
                int n = (int)GetNum(st.Dict, "N", 0);
                int first = (int)GetNum(st.Dict, "First", 0);
                if (n <= 0 || first <= 0 || first > data.Length) continue;

                PdfLexer head = new PdfLexer(data, 0);
                int[] nums = new int[n];
                int[] offs = new int[n];
                bool ok = true;
                for (int i = 0; i < n; i++)
                {
                    string a = head.ReadKeyword();
                    string b = head.ReadKeyword();
                    if (!PdfLexer.IsInt(a) || !PdfLexer.IsInt(b)) { ok = false; break; }
                    nums[i] = int.Parse(a, CultureInfo.InvariantCulture);
                    offs[i] = int.Parse(b, CultureInfo.InvariantCulture);
                }
                if (!ok) continue;

                for (int i = 0; i < n; i++)
                {
                    int p = first + offs[i];
                    if (p < 0 || p >= data.Length) continue;
                    // objects defined directly in the file win over packed ones
                    if (Objects.ContainsKey(nums[i]) && !(Objects[nums[i]] is PdfStream)) { }
                    PdfLexer lex = new PdfLexer(data, p);
                    PdfObj v = lex.ParseObject();
                    if (v != null && !Objects.ContainsKey(nums[i])) Objects[nums[i]] = v;
                }
            }
        }

        private void DetectEncryption()
        {
            byte[] needle = Encoding.ASCII.GetBytes("/Encrypt");
            Encrypted = IndexOf(Buf, needle, 0) >= 0;
        }

        // ---- page tree ----

        private void BuildPageList()
        {
            PdfDict catalog = null;
            foreach (KeyValuePair<int, PdfObj> kv in Objects)
            {
                PdfDict d = kv.Value as PdfDict;
                if (d == null) continue;
                PdfName t = d.Get("Type") as PdfName;
                if (t != null && t.Value == "Catalog") { catalog = d; break; }
            }

            if (catalog != null)
            {
                PdfDict pagesRoot = ResolveDict(catalog.Get("Pages"));
                if (pagesRoot != null)
                {
                    WalkPages(pagesRoot, new PdfDict(), 0, new HashSet<PdfDict>());
                    if (Pages.Count > 0) return;
                }
            }

            // No catalog, or a broken tree: take every /Type /Page in file order.
            List<int> keys = new List<int>(Objects.Keys);
            keys.Sort();
            for (int i = 0; i < keys.Count; i++)
            {
                PdfDict d = Objects[keys[i]] as PdfDict;
                if (d == null) continue;
                PdfName t = d.Get("Type") as PdfName;
                if (t != null && t.Value == "Page") Pages.Add(d);
            }
        }

        /// <summary>Resources, MediaBox and Rotate are inheritable: a page that
        /// does not state them uses its parent's. Push them down as we walk.</summary>
        private void WalkPages(PdfDict node, PdfDict inherited, int depth, HashSet<PdfDict> seen)
        {
            if (node == null || depth > 64 || Pages.Count > 5000) return;
            if (seen.Contains(node)) return;
            seen.Add(node);

            PdfDict inh = new PdfDict();
            foreach (KeyValuePair<string, PdfObj> kv in inherited.Map) inh.Map[kv.Key] = kv.Value;
            string[] inheritable = new string[] { "Resources", "MediaBox", "CropBox", "Rotate" };
            for (int i = 0; i < inheritable.Length; i++)
            {
                PdfObj v = node.Get(inheritable[i]);
                if (v != null) inh.Map[inheritable[i]] = v;
            }

            string type = GetName(node, "Type");
            PdfObj kidsObj = Resolve(node.Get("Kids"));
            PdfArray kids = kidsObj as PdfArray;

            if (kids != null && type != "Page")
            {
                for (int i = 0; i < kids.Items.Count; i++)
                    WalkPages(ResolveDict(kids.Items[i]), inh, depth + 1, seen);
                return;
            }

            PdfDict page = new PdfDict();
            foreach (KeyValuePair<string, PdfObj> kv in node.Map) page.Map[kv.Key] = kv.Value;
            foreach (KeyValuePair<string, PdfObj> kv in inh.Map)
                if (!page.Map.ContainsKey(kv.Key)) page.Map[kv.Key] = kv.Value;
            Pages.Add(page);
        }
    }

    // ---------------------------------------------------------------- fonts

    /// <summary>
    /// Everything we need from a font: how many bytes a character code takes,
    /// what Unicode it means, and how wide it is (so we know where the NEXT
    /// glyph lands, which is how word gaps are detected).
    /// </summary>
    public sealed class FontInfo
    {
        public Dictionary<int, string> ToUni;
        public Dictionary<int, double> Widths;   // /1000 text-space units
        public double DefaultWidth;
        public bool TwoByte;
        public bool Bold;
        public string BaseFont;

        public FontInfo()
        {
            ToUni = new Dictionary<int, string>();
            Widths = new Dictionary<int, double>();
            DefaultWidth = 0.5;
            BaseFont = "";
        }

        public string Decode(int code)
        {
            string s;
            if (ToUni.TryGetValue(code, out s)) return s;
            if (TwoByte) return "";
            if (code >= 32 && code < 127) return ((char)code).ToString();
            string w = Encodings.WinAnsi(code);
            if (w != null) return w;
            return "";
        }

        public double Width(int code)
        {
            double w;
            if (Widths.TryGetValue(code, out w)) return w;
            return DefaultWidth;
        }
    }

    public static class Encodings
    {
        /// <summary>The places WinAnsi differs from Latin-1 (the 0x80-0x9F band)
        /// plus the quotes and dashes that turn up constantly in invoice text.</summary>
        public static string WinAnsi(int code)
        {
            switch (code)
            {
                case 0x80: return "\u20AC"; case 0x82: return "\u201A";
                case 0x83: return "\u0192"; case 0x84: return "\u201E";
                case 0x85: return "\u2026"; case 0x86: return "\u2020";
                case 0x87: return "\u2021"; case 0x88: return "\u02C6";
                case 0x89: return "\u2030"; case 0x8A: return "\u0160";
                case 0x8B: return "\u2039"; case 0x8C: return "\u0152";
                case 0x8E: return "\u017D"; case 0x91: return "'";
                case 0x92: return "'";      case 0x93: return "\"";
                case 0x94: return "\"";     case 0x95: return "\u2022";
                case 0x96: return "-";      case 0x97: return "-";
                case 0x98: return "\u02DC"; case 0x99: return "\u2122";
                case 0x9A: return "\u0161"; case 0x9B: return "\u203A";
                case 0x9C: return "\u0153"; case 0x9E: return "\u017E";
                case 0x9F: return "\u0178";
            }
            if (code >= 0xA0 && code <= 0xFF) return ((char)code).ToString();
            return null;
        }

        /// <summary>Glyph name to text, for fonts that use /Differences.
        /// Covers the printable ASCII names plus uniXXXX and gNN forms.</summary>
        public static string GlyphName(string n)
        {
            if (string.IsNullOrEmpty(n)) return "";
            if (n.Length == 1) return n;

            if (n.Length >= 7 && n.Substring(0, 3) == "uni")
            {
                int v;
                if (int.TryParse(n.Substring(3, 4), NumberStyles.HexNumber,
                                 CultureInfo.InvariantCulture, out v)) return ((char)v).ToString();
            }
            if (n.Length >= 2 && n[0] == 'u' && n.Length <= 7)
            {
                int v;
                if (int.TryParse(n.Substring(1), NumberStyles.HexNumber,
                                 CultureInfo.InvariantCulture, out v) && v > 0 && v < 0x10000)
                    return ((char)v).ToString();
            }

            switch (n)
            {
                case "space": return " ";           case "exclam": return "!";
                case "quotedbl": return "\"";       case "numbersign": return "#";
                case "dollar": return "$";          case "percent": return "%";
                case "ampersand": return "&";       case "quotesingle": return "'";
                case "quoteright": return "'";      case "quoteleft": return "'";
                case "parenleft": return "(";       case "parenright": return ")";
                case "asterisk": return "*";        case "plus": return "+";
                case "comma": return ",";           case "hyphen": return "-";
                case "period": return ".";          case "slash": return "/";
                case "zero": return "0";            case "one": return "1";
                case "two": return "2";             case "three": return "3";
                case "four": return "4";            case "five": return "5";
                case "six": return "6";             case "seven": return "7";
                case "eight": return "8";           case "nine": return "9";
                case "colon": return ":";           case "semicolon": return ";";
                case "less": return "<";            case "equal": return "=";
                case "greater": return ">";         case "question": return "?";
                case "at": return "@";              case "bracketleft": return "[";
                case "backslash": return "\\";      case "bracketright": return "]";
                case "asciicircum": return "^";     case "underscore": return "_";
                case "grave": return "`";           case "braceleft": return "{";
                case "bar": return "|";             case "braceright": return "}";
                case "asciitilde": return "~";      case "quotedblleft": return "\"";
                case "quotedblright": return "\"";  case "endash": return "-";
                case "emdash": return "-";          case "bullet": return "\u2022";
                case "fi": return "fi";             case "fl": return "fl";
                case "sterling": return "\u00A3";   case "Euro": return "\u20AC";
                case "cent": return "\u00A2";       case "degree": return "\u00B0";
                case "numbersignsign": return "#";
            }
            return "";
        }
    }

    // ------------------------------------------------------------ the output

    /// <summary>One word, positioned on the page. X/Y are points from the
    /// TOP-LEFT of the page, because that is how a human reads an invoice.</summary>
    public sealed class PdfWord
    {
        public int Page;
        public string Text;
        public double X;
        public double Y;
        public double W;
        public double H;
        public double Size;
        public bool Bold;
    }

    public sealed class PdfPageInfo
    {
        public int Number;
        public double Width;
        public double Height;
    }

    public sealed class PdfTextResult
    {
        public List<PdfWord> Words;
        public List<PdfPageInfo> PageInfo;
        public bool Encrypted;
        public string Error;
        public PdfTextResult()
        {
            Words = new List<PdfWord>();
            PageInfo = new List<PdfPageInfo>();
        }
    }

    // ------------------------------------------------- the content interpreter

    internal sealed class Glyph
    {
        public string Text;
        public double X, Y, AdvX, Size;
        public bool Bold;
    }

    public static class PdfText
    {
        /// <summary>The whole job: a PDF path in, positioned words out.</summary>
        public static PdfTextResult Extract(string path)
        {
            PdfTextResult res = new PdfTextResult();
            PdfDoc doc;
            try { doc = PdfDoc.Load(path); }
            catch (Exception ex) { res.Error = ex.Message; return res; }

            res.Encrypted = doc.Encrypted;

            for (int pi = 0; pi < doc.Pages.Count; pi++)
            {
                PdfDict page = doc.Pages[pi];
                double x0 = 0, y0 = 0, x1 = 612, y1 = 792;
                PdfArray mb = doc.Resolve(page.Get("MediaBox")) as PdfArray;
                if (mb != null && mb.Items.Count >= 4)
                {
                    double[] v = new double[4];
                    bool ok = true;
                    for (int k = 0; k < 4; k++)
                    {
                        PdfNum n = doc.Resolve(mb.Items[k]) as PdfNum;
                        if (n == null) { ok = false; break; }
                        v[k] = n.Value;
                    }
                    if (ok)
                    {
                        x0 = Math.Min(v[0], v[2]); x1 = Math.Max(v[0], v[2]);
                        y0 = Math.Min(v[1], v[3]); y1 = Math.Max(v[1], v[3]);
                    }
                }
                int rot = ((int)doc.GetNum(page, "Rotate", 0) % 360 + 360) % 360;

                PdfPageInfo info = new PdfPageInfo();
                info.Number = pi + 1;
                info.Width = (rot == 90 || rot == 270) ? (y1 - y0) : (x1 - x0);
                info.Height = (rot == 90 || rot == 270) ? (x1 - x0) : (y1 - y0);
                res.PageInfo.Add(info);

                byte[] content = GetPageContent(doc, page);
                if (content.Length == 0) continue;

                List<Glyph> glyphs = new List<Glyph>();
                double[] baseCtm = new double[] { 1, 0, 0, 1, 0, 0 };
                try
                {
                    Run(doc, content, doc.ResolveDict(page.Get("Resources")),
                        baseCtm, glyphs, 0);
                }
                catch (Exception) { /* keep whatever this page yielded */ }

                BuildWords(glyphs, res.Words, pi + 1, x0, y0, x1, y1, rot);
            }
            return res;
        }

        private static byte[] GetPageContent(PdfDoc doc, PdfDict page)
        {
            PdfObj c = doc.Resolve(page.Get("Contents"));
            List<byte> all = new List<byte>();
            PdfStream st = c as PdfStream;
            if (st != null) all.AddRange(doc.GetStreamData(st));
            PdfArray arr = c as PdfArray;
            if (arr != null)
            {
                for (int i = 0; i < arr.Items.Count; i++)
                {
                    PdfStream s2 = doc.Resolve(arr.Items[i]) as PdfStream;
                    if (s2 == null) continue;
                    all.AddRange(doc.GetStreamData(s2));
                    all.Add(10);                     // streams must not fuse together
                }
            }
            return all.ToArray();
        }

        private static double[] Mul(double[] m, double[] n)
        {
            return new double[] {
                m[0]*n[0] + m[1]*n[2],
                m[0]*n[1] + m[1]*n[3],
                m[2]*n[0] + m[3]*n[2],
                m[2]*n[1] + m[3]*n[3],
                m[4]*n[0] + m[5]*n[2] + n[4],
                m[4]*n[1] + m[5]*n[3] + n[5]
            };
        }

        /// <summary>Interprets a content stream. Recurses into Form XObjects,
        /// which is where a surprising amount of invoice text actually lives.</summary>
        private static void Run(PdfDoc doc, byte[] content, PdfDict resources,
                                double[] ctm, List<Glyph> glyphs, int depth)
        {
            if (depth > 12) return;

            PdfDict fonts = doc.ResolveDict(resources == null ? null : resources.Get("Font"));
            PdfDict xobjs = doc.ResolveDict(resources == null ? null : resources.Get("XObject"));
            Dictionary<string, FontInfo> fontCache = new Dictionary<string, FontInfo>(StringComparer.Ordinal);

            List<double[]> ctmStack = new List<double[]>();
            double[] cur = (double[])ctm.Clone();

            double[] tm = new double[] { 1, 0, 0, 1, 0, 0 };
            double[] tlm = new double[] { 1, 0, 0, 1, 0, 0 };
            FontInfo font = null;
            double fs = 0, tc = 0, tw = 0, th = 1, tl = 0, ts = 0;

            List<PdfObj> stack = new List<PdfObj>();
            PdfLexer lex = new PdfLexer(content, 0);

            while (lex.Pos < content.Length)
            {
                int before = lex.Pos;
                PdfObj o = lex.ParseObject();
                if (o != null)
                {
                    if (stack.Count < 64) stack.Add(o);
                    continue;
                }
                string op = lex.LastKeyword;
                if (op == null)
                {
                    if (lex.Pos == before) lex.Pos++;
                    continue;
                }

                switch (op)
                {
                    case "q":
                        ctmStack.Add((double[])cur.Clone());
                        break;
                    case "Q":
                        if (ctmStack.Count > 0)
                        {
                            cur = ctmStack[ctmStack.Count - 1];
                            ctmStack.RemoveAt(ctmStack.Count - 1);
                        }
                        break;
                    case "cm":
                        if (stack.Count >= 6) cur = Mul(Nums(stack, 6), cur);
                        break;
                    case "BT":
                        tm = new double[] { 1, 0, 0, 1, 0, 0 };
                        tlm = (double[])tm.Clone();
                        break;
                    case "ET":
                        break;
                    case "Tf":
                        if (stack.Count >= 2)
                        {
                            PdfName fn = stack[stack.Count - 2] as PdfName;
                            PdfNum sz = stack[stack.Count - 1] as PdfNum;
                            if (sz != null) fs = sz.Value;
                            if (fn != null)
                            {
                                if (!fontCache.TryGetValue(fn.Value, out font))
                                {
                                    font = BuildFont(doc, doc.ResolveDict(
                                        fonts == null ? null : fonts.Get(fn.Value)));
                                    fontCache[fn.Value] = font;
                                }
                            }
                        }
                        break;
                    case "Td":
                        if (stack.Count >= 2)
                        {
                            double[] t = Nums(stack, 2);
                            tlm = Mul(new double[] { 1, 0, 0, 1, t[0], t[1] }, tlm);
                            tm = (double[])tlm.Clone();
                        }
                        break;
                    case "TD":
                        if (stack.Count >= 2)
                        {
                            double[] t = Nums(stack, 2);
                            tl = -t[1];
                            tlm = Mul(new double[] { 1, 0, 0, 1, t[0], t[1] }, tlm);
                            tm = (double[])tlm.Clone();
                        }
                        break;
                    case "Tm":
                        if (stack.Count >= 6)
                        {
                            tlm = Nums(stack, 6);
                            tm = (double[])tlm.Clone();
                        }
                        break;
                    case "T*":
                        tlm = Mul(new double[] { 1, 0, 0, 1, 0, -tl }, tlm);
                        tm = (double[])tlm.Clone();
                        break;
                    case "TL": if (stack.Count >= 1) tl = Num(stack, 1); break;
                    case "Tc": if (stack.Count >= 1) tc = Num(stack, 1); break;
                    case "Tw": if (stack.Count >= 1) tw = Num(stack, 1); break;
                    case "Tz": if (stack.Count >= 1) th = Num(stack, 1) / 100.0; break;
                    case "Ts": if (stack.Count >= 1) ts = Num(stack, 1); break;

                    case "Tj":
                        if (stack.Count >= 1)
                        {
                            PdfStr s = stack[stack.Count - 1] as PdfStr;
                            if (s != null)
                                Show(s.Bytes, font, ref tm, cur, fs, tc, tw, th, ts, glyphs);
                        }
                        break;

                    case "'":
                        tlm = Mul(new double[] { 1, 0, 0, 1, 0, -tl }, tlm);
                        tm = (double[])tlm.Clone();
                        if (stack.Count >= 1)
                        {
                            PdfStr s = stack[stack.Count - 1] as PdfStr;
                            if (s != null)
                                Show(s.Bytes, font, ref tm, cur, fs, tc, tw, th, ts, glyphs);
                        }
                        break;

                    case "\"":
                        if (stack.Count >= 3)
                        {
                            PdfNum aw = stack[stack.Count - 3] as PdfNum;
                            PdfNum ac = stack[stack.Count - 2] as PdfNum;
                            if (aw != null) tw = aw.Value;
                            if (ac != null) tc = ac.Value;
                            tlm = Mul(new double[] { 1, 0, 0, 1, 0, -tl }, tlm);
                            tm = (double[])tlm.Clone();
                            PdfStr s = stack[stack.Count - 1] as PdfStr;
                            if (s != null)
                                Show(s.Bytes, font, ref tm, cur, fs, tc, tw, th, ts, glyphs);
                        }
                        break;

                    case "TJ":
                        if (stack.Count >= 1)
                        {
                            PdfArray arr = stack[stack.Count - 1] as PdfArray;
                            if (arr != null)
                            {
                                for (int i = 0; i < arr.Items.Count; i++)
                                {
                                    PdfStr s = arr.Items[i] as PdfStr;
                                    if (s != null)
                                    {
                                        Show(s.Bytes, font, ref tm, cur, fs, tc, tw, th, ts, glyphs);
                                        continue;
                                    }
                                    PdfNum n = arr.Items[i] as PdfNum;
                                    if (n != null)
                                    {
                                        double dx = -n.Value / 1000.0 * fs * th;
                                        tm = Mul(new double[] { 1, 0, 0, 1, dx, 0 }, tm);
                                    }
                                }
                            }
                        }
                        break;

                    case "Do":
                        if (stack.Count >= 1 && xobjs != null)
                        {
                            PdfName xn = stack[stack.Count - 1] as PdfName;
                            if (xn != null)
                            {
                                PdfStream xs = doc.Resolve(xobjs.Get(xn.Value)) as PdfStream;
                                if (xs != null && doc.GetName(xs.Dict, "Subtype") == "Form")
                                {
                                    double[] fm = new double[] { 1, 0, 0, 1, 0, 0 };
                                    PdfArray ma = doc.Resolve(xs.Dict.Get("Matrix")) as PdfArray;
                                    if (ma != null && ma.Items.Count >= 6)
                                        for (int k = 0; k < 6; k++)
                                        {
                                            PdfNum n = doc.Resolve(ma.Items[k]) as PdfNum;
                                            if (n != null) fm[k] = n.Value;
                                        }
                                    PdfDict xr = doc.ResolveDict(xs.Dict.Get("Resources"));
                                    if (xr == null) xr = resources;
                                    Run(doc, doc.GetStreamData(xs), xr, Mul(fm, cur), glyphs, depth + 1);
                                }
                            }
                        }
                        break;

                    case "BI":
                        // Inline image: skip past its binary payload to the EI.
                        lex.Pos = SkipInlineImage(content, lex.Pos);
                        break;
                }
                stack.Clear();
            }
        }

        private static int SkipInlineImage(byte[] c, int pos)
        {
            int i = pos;
            while (i + 1 < c.Length)
            {
                if (c[i] == (byte)'I' && c[i + 1] == (byte)'D') { i += 2; break; }
                i++;
            }
            while (i + 1 < c.Length)
            {
                if (c[i] == (byte)'E' && c[i + 1] == (byte)'I' &&
                    (i == 0 || PdfLexer.IsWhite(c[i - 1])) &&
                    (i + 2 >= c.Length || !PdfLexer.IsRegular(c[i + 2])))
                    return i + 2;
                i++;
            }
            return c.Length;
        }

        private static double Num(List<PdfObj> st, int fromEnd)
        {
            PdfNum n = st[st.Count - fromEnd] as PdfNum;
            if (n == null) return 0;
            return n.Value;
        }

        private static double[] Nums(List<PdfObj> st, int count)
        {
            double[] r = new double[count];
            for (int i = 0; i < count; i++)
            {
                PdfNum n = st[st.Count - count + i] as PdfNum;
                r[i] = (n == null) ? 0 : n.Value;
            }
            return r;
        }

        /// <summary>Draws one string: each code becomes a glyph at the current
        /// point, then advances the text matrix by that glyph's width.</summary>
        private static void Show(byte[] bytes, FontInfo font, ref double[] tm, double[] ctm,
                                 double fs, double tc, double tw, double th, double ts,
                                 List<Glyph> glyphs)
        {
            if (bytes == null || bytes.Length == 0) return;
            if (font == null) font = new FontInfo();

            int step = font.TwoByte ? 2 : 1;
            for (int i = 0; i + step <= bytes.Length; i += step)
            {
                int code = (step == 2) ? ((bytes[i] << 8) | bytes[i + 1]) : bytes[i];

                double w0 = font.Width(code);
                double[] trm = Mul(new double[] { fs * th, 0, 0, fs, 0, ts }, Mul(tm, ctm));

                double scaleY = Math.Sqrt(trm[2] * trm[2] + trm[3] * trm[3]);
                double scaleX = Math.Sqrt(trm[0] * trm[0] + trm[1] * trm[1]);

                string txt = font.Decode(code);
                bool isSpaceCode = (step == 1 && code == 32);

                double adv = (w0 * fs + tc + (isSpaceCode ? tw : 0)) * th;

                Glyph g = new Glyph();
                g.Text = txt;
                g.X = trm[4];
                g.Y = trm[5];
                g.Size = (scaleY > 0.0001) ? scaleY : Math.Abs(fs);
                g.AdvX = adv * ((scaleX > 0.0001 && fs != 0) ? (scaleX / Math.Abs(fs * th)) : 1);
                g.Bold = font.Bold;
                if (isSpaceCode && txt.Length == 0) g.Text = " ";
                glyphs.Add(g);

                tm = Mul(new double[] { 1, 0, 0, 1, adv, 0 }, tm);
            }
        }

        /// <summary>
        /// Glyphs to words. A word ends at a space, at a vertical step, or at a
        /// horizontal gap wider than a fraction of the font size -- which is how
        /// "Invoice Date" stays two words but "Total:" stays one.
        /// </summary>
        private static void BuildWords(List<Glyph> glyphs, List<PdfWord> outWords, int pageNo,
                                       double x0, double y0, double x1, double y1, int rot)
        {
            StringBuilder sb = new StringBuilder();
            double wx = 0, wy = 0, wEnd = 0, wSize = 0;
            bool wBold = false;
            bool open = false;

            for (int i = 0; i <= glyphs.Count; i++)
            {
                Glyph g = (i < glyphs.Count) ? glyphs[i] : null;

                bool breakHere = (g == null);
                if (!breakHere && open)
                {
                    if (Math.Abs(g.Y - wy) > Math.Max(1.0, wSize * 0.35)) breakHere = true;
                    else if (g.X - wEnd > Math.Max(0.6, wSize * 0.22)) breakHere = true;
                    else if (g.X < wEnd - Math.Max(1.0, wSize * 1.5)) breakHere = true;   // jumped back
                }

                if (breakHere && open)
                {
                    string t = sb.ToString().Trim();
                    if (t.Length > 0)
                    {
                        PdfWord w = new PdfWord();
                        w.Page = pageNo;
                        w.Text = t;
                        w.Size = wSize;
                        w.Bold = wBold;
                        double left = wx, right = wEnd, baseline = wy;
                        MapCoords(left, right, baseline, wSize, x0, y0, x1, y1, rot, w);
                        outWords.Add(w);
                    }
                    sb.Length = 0;
                    open = false;
                }

                if (g == null) break;

                if (g.Text == " " || g.Text == "\t")
                {
                    if (open)
                    {
                        string t = sb.ToString().Trim();
                        if (t.Length > 0)
                        {
                            PdfWord w = new PdfWord();
                            w.Page = pageNo;
                            w.Text = t;
                            w.Size = wSize;
                            w.Bold = wBold;
                            MapCoords(wx, wEnd, wy, wSize, x0, y0, x1, y1, rot, w);
                            outWords.Add(w);
                        }
                        sb.Length = 0;
                        open = false;
                    }
                    continue;
                }
                if (g.Text.Length == 0) { if (open) wEnd = g.X + g.AdvX; continue; }

                if (!open)
                {
                    wx = g.X; wy = g.Y; wSize = g.Size; wBold = g.Bold;
                    open = true;
                }
                if (g.Size > wSize) wSize = g.Size;
                if (g.Bold) wBold = true;
                sb.Append(g.Text);
                wEnd = g.X + g.AdvX;
            }
        }

        /// <summary>PDF space (origin bottom-left, y up) to reading space
        /// (origin top-left, y down), honouring /Rotate.</summary>
        private static void MapCoords(double left, double right, double baseline, double size,
                                      double x0, double y0, double x1, double y1, int rot,
                                      PdfWord w)
        {
            double wdt = Math.Max(0, right - left);
            double asc = size * 0.78;
            double dsc = size * 0.22;

            switch (rot)
            {
                case 90:
                    w.X = baseline - y0 - dsc;
                    w.Y = left - x0;
                    w.W = size;
                    w.H = wdt;
                    break;
                case 180:
                    w.X = x1 - right;
                    w.Y = baseline - y0 - dsc;
                    w.W = wdt;
                    w.H = size;
                    break;
                case 270:
                    w.X = y1 - baseline - asc;
                    w.Y = x1 - right;
                    w.W = size;
                    w.H = wdt;
                    break;
                default:
                    w.X = left - x0;
                    w.Y = y1 - baseline - asc;
                    w.W = wdt;
                    w.H = size;
                    break;
            }
        }

        // ---- font building ----

        private static FontInfo BuildFont(PdfDoc doc, PdfDict fd)
        {
            FontInfo fi = new FontInfo();
            if (fd == null) return fi;

            string bf = doc.GetName(fd, "BaseFont");
            if (bf != null)
            {
                fi.BaseFont = bf;
                string up = bf.ToUpperInvariant();
                fi.Bold = up.Contains("BOLD") || up.Contains("BLACK") || up.Contains("HEAVY");
            }

            string subtype = doc.GetName(fd, "Subtype");
            PdfDict desc = null;

            if (subtype == "Type0")
            {
                string enc = doc.GetName(fd, "Encoding");
                fi.TwoByte = true;                       // Identity-H and friends
                if (enc != null && enc.Contains("UCS2")) fi.TwoByte = true;

                PdfArray df = doc.Resolve(fd.Get("DescendantFonts")) as PdfArray;
                if (df != null && df.Items.Count > 0)
                {
                    PdfDict d0 = doc.ResolveDict(df.Items[0]);
                    if (d0 != null)
                    {
                        fi.DefaultWidth = doc.GetNum(d0, "DW", 1000) / 1000.0;
                        ReadCidWidths(doc, doc.Resolve(d0.Get("W")) as PdfArray, fi);
                        desc = doc.ResolveDict(d0.Get("FontDescriptor"));
                    }
                }
            }
            else
            {
                int first = (int)doc.GetNum(fd, "FirstChar", 0);
                PdfArray wArr = doc.Resolve(fd.Get("Widths")) as PdfArray;
                if (wArr != null)
                {
                    for (int i = 0; i < wArr.Items.Count; i++)
                    {
                        PdfNum n = doc.Resolve(wArr.Items[i]) as PdfNum;
                        if (n != null) fi.Widths[first + i] = n.Value / 1000.0;
                    }
                }
                desc = doc.ResolveDict(fd.Get("FontDescriptor"));
                if (wArr == null) ApplyBuiltinWidths(fi);

                // /Encoding may remap individual codes
                PdfObj encObj = doc.Resolve(fd.Get("Encoding"));
                PdfDict encDict = encObj as PdfDict;
                if (encDict != null)
                {
                    PdfArray diff = doc.Resolve(encDict.Get("Differences")) as PdfArray;
                    if (diff != null)
                    {
                        int code = 0;
                        for (int i = 0; i < diff.Items.Count; i++)
                        {
                            PdfNum n = diff.Items[i] as PdfNum;
                            if (n != null) { code = (int)n.Value; continue; }
                            PdfName gn = diff.Items[i] as PdfName;
                            if (gn != null)
                            {
                                string t = Encodings.GlyphName(gn.Value);
                                if (t.Length > 0) fi.ToUni[code] = t;
                                code++;
                            }
                        }
                    }
                }
            }

            if (desc != null)
            {
                double flags = doc.GetNum(desc, "Flags", 0);
                if (((int)flags & 0x40000) != 0) fi.Bold = true;         // ForceBold
                double mw = doc.GetNum(desc, "MissingWidth", -1);
                if (mw >= 0 && fi.Widths.Count == 0) fi.DefaultWidth = mw / 1000.0;
                double sw = doc.GetNum(desc, "StemV", 0);
                if (sw >= 120) fi.Bold = true;
            }

            // /ToUnicode wins over everything: it is the font's own answer.
            PdfStream tu = doc.Resolve(fd.Get("ToUnicode")) as PdfStream;
            if (tu != null) ParseToUnicode(doc.GetStreamData(tu), fi);

            return fi;
        }

        private static void ReadCidWidths(PdfDoc doc, PdfArray w, FontInfo fi)
        {
            if (w == null) return;
            int i = 0;
            while (i < w.Items.Count)
            {
                PdfNum a = doc.Resolve(w.Items[i]) as PdfNum;
                if (a == null) { i++; continue; }
                if (i + 1 >= w.Items.Count) break;

                PdfObj nxt = doc.Resolve(w.Items[i + 1]);
                PdfArray list = nxt as PdfArray;
                if (list != null)
                {
                    int start = (int)a.Value;
                    for (int k = 0; k < list.Items.Count; k++)
                    {
                        PdfNum n = doc.Resolve(list.Items[k]) as PdfNum;
                        if (n != null) fi.Widths[start + k] = n.Value / 1000.0;
                    }
                    i += 2;
                    continue;
                }
                PdfNum b = nxt as PdfNum;
                if (b != null && i + 2 < w.Items.Count)
                {
                    PdfNum c = doc.Resolve(w.Items[i + 2]) as PdfNum;
                    if (c != null)
                    {
                        int lo = (int)a.Value, hi = (int)b.Value;
                        if (hi - lo < 65536)
                            for (int k = lo; k <= hi; k++) fi.Widths[k] = c.Value / 1000.0;
                    }
                    i += 3;
                    continue;
                }
                i++;
            }
        }

        /// <summary>The base-14 fonts carry no /Widths. Helvetica-ish metrics are
        /// close enough for gap detection, which is all widths are used for.</summary>
        private static void ApplyBuiltinWidths(FontInfo fi)
        {
            string bf = fi.BaseFont.ToUpperInvariant();
            if (bf.Contains("COURIER") || bf.Contains("MONO"))
            {
                fi.DefaultWidth = 0.6;
                return;
            }
            fi.DefaultWidth = 0.5;
            string narrow = "ijlt.,:;'`|!I()[]{}/\\ ";
            string wide = "mMWw@%";
            for (int c = 32; c < 127; c++)
            {
                char ch = (char)c;
                double w = 0.556;
                if (narrow.IndexOf(ch) >= 0) w = 0.28;
                else if (wide.IndexOf(ch) >= 0) w = 0.85;
                else if (ch >= 'A' && ch <= 'Z') w = 0.68;
                else if (ch >= '0' && ch <= '9') w = 0.556;
                else if (ch >= 'a' && ch <= 'z') w = 0.55;
                fi.Widths[c] = w;
            }
            fi.Widths[32] = 0.278;
        }

        /// <summary>Parses a /ToUnicode CMap: bfchar maps single codes, bfrange
        /// maps spans. Destinations are UTF-16BE.</summary>
        private static void ParseToUnicode(byte[] data, FontInfo fi)
        {
            if (data == null || data.Length == 0) return;
            PdfLexer lex = new PdfLexer(data, 0);
            List<PdfObj> stack = new List<PdfObj>();
            int maxSrcLen = 1;

            while (lex.Pos < data.Length)
            {
                int before = lex.Pos;
                PdfObj o = lex.ParseObject();
                if (o != null)
                {
                    if (stack.Count < 400) stack.Add(o);
                    continue;
                }
                string kw = lex.LastKeyword;
                if (kw == null) { if (lex.Pos == before) lex.Pos++; continue; }

                if (kw == "endbfchar")
                {
                    for (int i = 0; i + 1 < stack.Count; i += 2)
                    {
                        PdfStr src = stack[i] as PdfStr;
                        PdfStr dst = stack[i + 1] as PdfStr;
                        if (src == null || dst == null) continue;
                        if (src.Bytes.Length > maxSrcLen) maxSrcLen = src.Bytes.Length;
                        fi.ToUni[CodeOf(src.Bytes)] = Utf16Be(dst.Bytes);
                    }
                }
                else if (kw == "endbfrange")
                {
                    for (int i = 0; i + 2 < stack.Count; i += 3)
                    {
                        PdfStr lo = stack[i] as PdfStr;
                        PdfStr hi = stack[i + 1] as PdfStr;
                        if (lo == null || hi == null) continue;
                        if (lo.Bytes.Length > maxSrcLen) maxSrcLen = lo.Bytes.Length;
                        int a = CodeOf(lo.Bytes), b = CodeOf(hi.Bytes);
                        if (b < a || b - a > 65535) continue;

                        PdfArray arr = stack[i + 2] as PdfArray;
                        if (arr != null)
                        {
                            for (int k = 0; k <= b - a && k < arr.Items.Count; k++)
                            {
                                PdfStr d = arr.Items[k] as PdfStr;
                                if (d != null) fi.ToUni[a + k] = Utf16Be(d.Bytes);
                            }
                            continue;
                        }
                        PdfStr dst = stack[i + 2] as PdfStr;
                        if (dst == null) continue;
                        string baseStr = Utf16Be(dst.Bytes);
                        if (baseStr.Length == 0) continue;
                        int baseCp = baseStr[baseStr.Length - 1];
                        string prefix = baseStr.Substring(0, baseStr.Length - 1);
                        for (int k = 0; k <= b - a; k++)
                            fi.ToUni[a + k] = prefix + ((char)(baseCp + k)).ToString();
                    }
                }
                else if (kw == "endcodespacerange")
                {
                    for (int i = 0; i < stack.Count; i++)
                    {
                        PdfStr s = stack[i] as PdfStr;
                        if (s != null && s.Bytes.Length > maxSrcLen) maxSrcLen = s.Bytes.Length;
                    }
                }

                if (kw == "endbfchar" || kw == "endbfrange" || kw == "endcodespacerange" ||
                    kw == "beginbfchar" || kw == "beginbfrange" || kw == "begincodespacerange")
                    stack.Clear();
                if (stack.Count > 380) stack.Clear();
            }

            if (maxSrcLen >= 2) fi.TwoByte = true;
        }

        private static int CodeOf(byte[] b)
        {
            int v = 0;
            for (int i = 0; i < b.Length && i < 4; i++) v = (v << 8) | b[i];
            return v;
        }

        private static string Utf16Be(byte[] b)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i + 1 < b.Length; i += 2)
            {
                int cp = (b[i] << 8) | b[i + 1];
                sb.Append((char)cp);
            }
            if (b.Length == 1) sb.Append((char)b[0]);
            return sb.ToString();
        }
    }
}
