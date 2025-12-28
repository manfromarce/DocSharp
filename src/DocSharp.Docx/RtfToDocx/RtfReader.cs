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
                                stack.Peek().Tokens.Add(new RtfControlWord(sym) { HasValue = false });
                                break;
                        }
                    }
                    groupJustOpened = false;
                    continue;
                }

                // read letters (and, for destination names, digits) of control word
                var nameSb = new StringBuilder();
                nameSb.Append(next);
                while (true)
                {
                    int p = PeekChar();
                    if (p == -1) break;
                    char pc = (char)p;
                    // if this is a destination name (we previously saw \* right after '{'),
                    // include trailing digits as part of the name (e.g. pnseclvl1)
                    if (IsEnglishLetter(pc) || (pendingDestinationMarker && groupJustOpened && char.IsDigit(pc)))
                        nameSb.Append((char)ReadChar());
                    else
                        break;
                }
                string name = nameSb.ToString();
                // optional numeric parameter (signed) — skip parsing when this is a destination name
                bool hasNumber = false;
                int sign = 1;
                int value = 0;
                if (!(pendingDestinationMarker && groupJustOpened))
                {
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
                }

                bool delimitedBySpace = false;
                int p3 = PeekChar();
                if (p3 != -1 && (char)p3 == ' ') { delimitedBySpace = true; ReadChar(); }

                var cw = new RtfControlWord(name) { HasValue = hasNumber, Value = hasNumber ? (int?)value : null, DelimitedBySpace = delimitedBySpace };

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

                if (pendingDestinationMarker && groupJustOpened)
                {
                    // convert the recently created group into a RtfDestination with this name
                    var old = stack.Pop();
                    var parent = stack.Peek();
                    var dest = new RtfDestination(name);
                    // move any tokens that might already exist inside the group
                    dest.Tokens.AddRange(old.Tokens);
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
                // Ignore unless in \binN
            }
            else
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
