using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace DocSharp.Rtf;

internal static class RtfReader
{
#if !NETFRAMEWORK
    static RtfReader()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
#endif

    private static readonly string[] destinations = new[]
    {
        "aftncn",
        "aftnsep",
        "aftnsepc",
        "annotation",
        "atnauthor",
        "atndate",
        "atnicn",
        "atnparent",
        "atnref",
        "atntime",
        "atrfend",
        "atrfstart",
        "author",
        "background",
        "bkmkend",
        "bkmkstart",
        "blipuid",
        "buptim",
        "category",
        "colorschememapping",
        "colortbl",
        "comment",
        "company",
        "creatim",
        "datafield",
        "datastore",
        "defchp",
        "defpap",
        "do",
        "doccomm",
        "docvar",
        "dptxbtext",
        "ebcend",
        "ebcstart",
        "factoidname",
        "falt",
        "fchars",
        "ffdeftext",
        "ffentrymcr",
        "ffexitmcr",
        "ffformat",
        "ffhelptext",
        "ffl",
        "ffname",
        "ffstattext",
        "field",
        "file",
        "filetbl",
        "fldinst",
        "fldrslt",
        "fldtype",
        "fname",
        "fontemb",
        "fontfile",
        "fonttbl",
        "footer",
        "footerf",
        "footerl",
        "footerr",
        "footnote",
        "formfield",
        "ftncn",
        "ftnsep",
        "ftnsepc",
        "g",
        "generator",
        "gridtbl",
        "header",
        "headerf",
        "headerl",
        "headerr",
        "hl",
        "hlfr",
        "hlinkbase",
        "hlloc",
        "hlsrc",
        "hsv",
        "htmltag",
        "info",
        "keycode",
        "keywords",
        "latentstyles",
        "levelnumbers",
        "leveltext",
        "lfolevel",
        "linkval",
        "list",
        "listlevel",
        "listname",
        "listoverride",
        "listoverridetable",
        "listpicture",
        "liststylename",
        "listtable",
        "listtext",
        "lsdlockedexcept",
        "macc",
        "maccPr",
        "mailmerge",
        "maln",
        "malnScr",
        "manager",
        "margPr",
        "mbar",
        "mbarPr",
        "mbaseJc",
        "mbegChr",
        "mborderBox",
        "mborderBoxPr",
        "mbox",
        "mboxPr",
        "mchr",
        "mcount",
        "mctrlPr",
        "md",
        "mdeg",
        "mdegHide",
        "mden",
        "mdiff",
        "mdPr",
        "me",
        "mendChr",
        "meqArr",
        "meqArrPr",
        "mf",
        "mfName",
        "mfPr",
        "mfunc",
        "mfuncPr",
        "mgroupChr",
        "mgroupChrPr",
        "mgrow",
        "mhideBot",
        "mhideLeft",
        "mhideRight",
        "mhideTop",
        "mhtmltag",
        "mlim",
        "mlimloc",
        "mlimlow",
        "mlimlowPr",
        "mlimupp",
        "mlimuppPr",
        "mm",
        "mmaddfieldname",
        "mmath",
        "mmathPict",
        "mmathPr",
        "mmaxdist",
        "mmc",
        "mmcJc",
        "mmconnectstr",
        "mmconnectstrdata",
        "mmcPr",
        "mmcs",
        "mmdatasource",
        "mmheadersource",
        "mmmailsubject",
        "mmodso",
        "mmodsofilter",
        "mmodsofldmpdata",
        "mmodsomappedname",
        "mmodsoname",
        "mmodsorecipdata",
        "mmodsosort",
        "mmodsosrc",
        "mmodsotable",
        "mmodsoudl",
        "mmodsoudldata",
        "mmodsouniquetag",
        "mmPr",
        "mmquery",
        "mmr",
        "mnaryPr",
        "mnoBreak",
        "mnum",
        "mobjDist",
        "moMath",
        "moMathPara",
        "moMathParaPr",
        "mopEmu",
        "mphant",
        "mphantPr",
        "mplcHide",
        "mpos",
        "mr",
        "mrad",
        "mradPr",
        "mrPr",
        "msepChr",
        "mshow",
        "mshp",
        "mspre",
        "msPrePr",
        "msSub",
        "msSubPr",
        "msSubSup",
        "msSubSupPr",
        "msSup",
        "msSupPr",
        "mstrikeBLTR",
        "mstrikeH",
        "mstrikeTLBR",
        "mstrikeV",
        "msub",
        "msubHide",
        "msup",
        "msupHide",
        "mtransp",
        "mtype",
        "mvertJc",
        "mvfmf",
        "mvfml",
        "mvtof",
        "mvtol",
        "mzeroAsc",
        "mzeroDesc",
        "mzeroWid",
        "nesttableprops",
        "nextfile",
        "nonesttables",
        "objalias",
        "objclass",
        "objdata",
        "object",
        "objname",
        "objsect",
        "objtime",
        "oldcprops",
        "oldpprops",
        "oldsprops",
        "oldtprops",
        "oleclsid",
        "operator",
        "panose",
        "password",
        "passwordhash",
        "pgp",
        "pgptbl",
        "picprop",
        "pict",
        "pn",
        "pnseclvl", // special case: it's followed by a numeric parameter
        "pntext",
        "pntxta",
        "pntxtb",
        "printim",
        "private",
        "propname",
        "protend",
        "protstart",
        "protusertbl",
        "pxe",
        "result",
        "revtbl",
        "revtim",
        "rsidtbl",
        // "rtf",
        "rxe",
        "shp",
        "shpgrp",
        "shpinst",
        "shppict",
        "shprslt",
        "shptxt",
        "sn",
        "sp",
        "staticval",
        "stylesheet",
        "subject",
        "sv",
        "svb",
        "tc",
        "template",
        "themedata",
        "title",
        "txe",
        "ud",
        "upr",
        "userprops",
        "wgrffmtfilter",
        "windowcaption",
        "writereservation",
        "writereservhash",
        "xe",
        "xform",
        "xmlattrname",
        "xmlattrvalue",
        "xmlclose",
        "xmlname",
        "xmlnstbl",
        "xmlopen"
    };

    public static RtfDocument ReadRtf(TextReader reader)
    {
        var doc = new RtfDocument();
        var stack = new Stack<RtfGroup>();
        stack.Push(doc.Root);

        bool groupJustOpened = false;
        bool pendingDestinationMarker = false;

        var textBuf = new StringBuilder();

        void FlushText()
        {
            if (textBuf.Length > 0)
            {
                stack.Peek().Tokens.Add(new RtfText(textBuf.ToString()));
                textBuf.Clear();
            }
        }

        int pushback = -1;

        int ReadChar()
        {
            if (pushback != -1)
            {
                int t = pushback;
                pushback = -1;
                return t;
            }
            return reader.Read();
        }

        int PeekChar()
        {
            if (pushback != -1) return pushback;
            pushback = reader.Read();
            return pushback;
        }

        int r;
        while ((r = ReadChar()) != -1)
        {
            char c = (char)r;
            if (c == '{')
            {
                FlushText();
                var g = new RtfGroup();
                stack.Peek().Tokens.Add(g);
                stack.Push(g);
                groupJustOpened = true;
                pendingDestinationMarker = false;
            }
            else if (c == '}')
            {
                FlushText();
                if (stack.Count > 1) stack.Pop();
                groupJustOpened = false;
                pendingDestinationMarker = false;
            }
            else if (c == '\\')
            {
                FlushText();
                int pk = PeekChar();
                if (pk == -1) break;
                char next = (char)ReadChar();

                if (next == '\'')
                {
                    // hex escape: two hex digits
                    int h1 = ReadChar();
                    int h2 = ReadChar();
                    if (h1 != -1 && h2 != -1)
                    {
                        string hex = new string(new[] { (char)h1, (char)h2 });
                        if (int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out int v))
                        {
                            stack.Peek().Tokens.Add(new RtfText(((char)v).ToString()));
                        }
                    }
                    groupJustOpened = false;
                    continue;
                }

                if (!IsEnglishLetter(next))
                {
                    string sym = next.ToString();
                    if (sym == "*")
                    {
                        if (groupJustOpened)
                        {
                            // mark that the next control word is the destination name
                            pendingDestinationMarker = true;
                            // keep groupJustOpened = true so the following control word
                            // can detect and convert this group into a RtfDestination
                            continue;
                        }
                    }
                    else if (sym == "\\" || sym == "{" || sym == "}")
                    {
                        stack.Peek().Tokens.Add(new RtfText(sym));
                    }
                    else
                    {
                        // map single-character control symbols to text where appropriate
                        switch (sym)
                        {
                            case "~": // non-breaking space
                                stack.Peek().Tokens.Add(new RtfText("\u00A0"));
                                break;
                            case "-": // soft hyphen
                                stack.Peek().Tokens.Add(new RtfText("\u00AD"));
                                break;
                            case "_": // non-breaking hyphen
                                stack.Peek().Tokens.Add(new RtfText("\u2011"));
                                break;
                            default:
                                stack.Peek().Tokens.Add(new RtfControlWord(sym));
                                break;
                        }
                    }
                    groupJustOpened = false;
                    continue;
                }

                // read letters of control word / destination
                var nameSb = new StringBuilder();
                nameSb.Append(next);
                while (true)
                {
                    int p = PeekChar();
                    if (p == -1) break;
                    char pc = (char)p;
                    if (IsEnglishLetter(pc))
                        nameSb.Append((char)ReadChar());
                    else
                        break;
                }
                string name = nameSb.ToString();
                // optional numeric parameter (signed) — skip parsing when this is a destination name
                bool hasNumber = false;
                int sign = 1;
                int value = 0;
                
                int p2 = PeekChar();
                if (p2 != -1)
                {
                    char pc2 = (char)p2;
                    if (pc2 == '-') { ReadChar(); sign = -1; p2 = PeekChar(); }
                    if (p2 != -1 && char.IsDigit((char)p2))
                    {
                        hasNumber = true;
                        int acc = 0;
                        while (true)
                        {
                            int d = PeekChar(); if (d == -1) break;
                            char dc = (char)d; if (!char.IsDigit(dc)) break;
                            acc = acc * 10 + (dc - '0'); ReadChar();
                        }
                        value = acc * sign;
                    }
                }

                bool delimitedBySpace = false;
                int p3 = PeekChar();
                if (p3 != -1 && (char)p3 == ' ') { delimitedBySpace = true; ReadChar(); }

                var cw = new RtfControlWord(name) { Value = hasNumber ? (int?)value : null, DelimitedBySpace = delimitedBySpace };

                // map certain control words to text tokens (spaces, dashes, quotes)
                bool handledAsText = false;
                switch (name.ToLowerInvariant())
                {
                    case "enspace": // U+2002
                        stack.Peek().Tokens.Add(new RtfText("\u2002")); handledAsText = true; break;
                    case "emspace": // U+2003
                        stack.Peek().Tokens.Add(new RtfText("\u2003")); handledAsText = true; break;
                    case "qmspace": // U+2005 (four-per-em)
                        stack.Peek().Tokens.Add(new RtfText("\u2005")); handledAsText = true; break;
                    case "endash": // U+2013
                        stack.Peek().Tokens.Add(new RtfText("\u2013")); handledAsText = true; break;
                    case "emdash": // U+2014
                        stack.Peek().Tokens.Add(new RtfText("\u2014")); handledAsText = true; break;
                    case "lquote": // U+2018
                        stack.Peek().Tokens.Add(new RtfText("\u2018")); handledAsText = true; break;
                    case "rquote": // U+2019
                        stack.Peek().Tokens.Add(new RtfText("\u2019")); handledAsText = true; break;
                    case "ldblquote": // U+201C
                        stack.Peek().Tokens.Add(new RtfText("\u201C")); handledAsText = true; break;
                    case "rdblquote": // U+201D
                        stack.Peek().Tokens.Add(new RtfText("\u201D")); handledAsText = true; break;
                    case "bullet": // U+2022 (•)
                        stack.Peek().Tokens.Add(new RtfText("\u2022")); handledAsText = true; break;
                }
                // Better handled as control words: \line, \tab, \page, \column, \par, \sect.
                
                // Also TODO: \, { and } should be considered text when escaped by a previous \ 
                // (not recommended by the specification but might be found in existing files)

                if (handledAsText)
                {
                    // if this was a destination marker situation, still need to handle it below
                    if (pendingDestinationMarker && groupJustOpened)
                    {
                        // in destination case, the name is meaningful; fall through to destination handling
                    }
                    else
                    {
                        groupJustOpened = false;
                        continue;
                    }
                }

                // If this group was just opened and the control word names a known destination,
                // convert the recently created group into an RtfDestination. This should happen
                // both for destinations preceded by '*' (pendingDestinationMarker) and for
                // plain destination names (e.g. \colortbl).
                bool isDest = groupJustOpened && (pendingDestinationMarker || Array.IndexOf(destinations, name.ToLowerInvariant()) >= 0);
                if (isDest)
                {
                    var old = stack.Pop();
                    var parent = stack.Peek();
                    var dest = new RtfDestination(name, pendingDestinationMarker);
                    // move any tokens that might already exist inside the group
                    dest.Tokens.AddRange(old.Tokens);
                    // preserve numeric parameter when present (some destinations like pnseclvl carry a value)
                    dest.Value = cw.Value;
                    // replace the last token in parent with the destination
                    int idx = parent.Tokens.Count - 1;
                    if (idx >= 0)
                    {
                        parent.Tokens[idx] = dest;
                    }
                    else
                    {
                        parent.Tokens.Add(dest);
                    }
                    stack.Push(dest);
                    pendingDestinationMarker = false;
                }
                else
                {
                    stack.Peek().Tokens.Add(cw);
                }

                groupJustOpened = false;
            }
            else if (c == '\r' || c == '\n')
            {
                // Ignore unless in \binN (or picture ^)
            }
            // else if (c != '\0')
            else if (!char.IsControl(c))
            {    
                textBuf.Append(c);
                groupJustOpened = false;             
            }
        }

        if (textBuf.Length > 0) 
            stack.Peek().Tokens.Add(new RtfText(textBuf.ToString()));

        return doc;
    }

    private static bool IsEnglishLetter(char c)
    {
#if NETFRAMEWORK
        return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
#else
        return char.IsAsciiLetter(c);
#endif
    }    
}
