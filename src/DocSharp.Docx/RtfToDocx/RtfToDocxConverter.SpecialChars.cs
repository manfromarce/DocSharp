using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
    private bool ProcessSpecialCharControlWord(RtfControlWord cw, FormattingState runState)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch(name)
        {
            // TODO: use the current culture specified in RTF for the fallback string of chdate and chtime
            case "chdate": 
                CreateField("date", DateTime.Now.ToShortDateString());
                break;
            case "chtime":
                CreateField("time", DateTime.Now.ToShortTimeString());
                break;

            // Note: these are formatted by Word using the English culture
            case "chdpl": 
                CreateField("date \\@ \"dddd, MMMM d, yyyy\"", DateTime.Now.ToString("dddd, MMMM d, yyyy"));
                break;
            case "chdpa": 
                CreateField("date \\@ \"ddd, MMM d, yyyy\"", DateTime.Now.ToString("ddd, MMM d, yyyy"));
                break;

            case "sectnum": // TODO: keep track of the current section number and write it as fallback
                CreateSimpleField(" SECTION \\* MERGEFORMAT ", "1");
                break;
            // TODO: create comment (followed by the content group)
            // case "chatn": 
            //     break;
            case "chpgn": 
                AddRun().Append(new PageNumber());
                break;
            case "tab":
                EnsureRun().Append(new TabChar());
                break;
            case "uc":
                // Number of ANSI characters to skip after a following \uN control word
                if (cw.HasValue)
                    runState.Uc = Math.Max(0, cw.Value!.Value);
                else
                    runState.Uc = 1;
                break;
            case "u":
                if (cw.HasValue)
                {
                    ProcessUnicode(cw.Value!.Value, runState);
                }
                break;
        }
        return false;
    }

    private void ProcessUnicode(int charCode, FormattingState runState, StringBuilder? sb = null)
    {
        if (charCode < 0)
        {
            // Unicode values greater than 32767 are expressed as negative numbers.
            // For example, U+F020 would be \u-4064 in RTF: 
            // sum 65536 to get 61472.
            charCode += 65536;
        }
        var pending = runState.PendingHighSurrogate;
        if (charCode >= 0xD800 && charCode <= 0xDBFF)
        {
            // High surrogate: buffer until low surrogate arrives
            runState.PendingHighSurrogate = charCode;
        }
        else if (charCode >= 0xDC00 && charCode <= 0xDFFF && pending.HasValue)
        {
            // Low surrogate: combine with pending high if present
            var high = (char)pending.Value;
            var low = (char)charCode;
            HandleText(new string(new char[] { high, low }), sb);
            runState.PendingHighSurrogate = null;
        }
        else if (charCode >= 0 && charCode <= 0x10FFFF)
        {
            // Normal codepoint: if a pending high exists, it didn't pair
            if (pending.HasValue)
            {
                runState.PendingHighSurrogate = null;
            }
            string s = char.ConvertFromUtf32(charCode);
            HandleText(s, sb);
        }
        // After emitting the Unicode character, the following ANSI characters should be ignored based on the "uc" number. 
        // Track how many to skip on the formatting state so subsequent text tokens can consume them.
        runState.PendingAnsiSkip = Math.Max(runState.Uc, 0);
    }
}