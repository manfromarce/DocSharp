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
            // TODO: create comments and footnotes/endnotes (followed by the content group)
            // case "chatn": 
            //     break;
            // case "chftn": 
            //     break;
            case "chftnsep": 
                EnsureRun();
                currentRun!.Append(new SeparatorMark());
                break;
            case "chftnsepc":
                EnsureRun();
                currentRun!.Append(new ContinuationSeparatorMark());
                break;
            case "chpgn": 
                EnsureRun();
                currentRun!.Append(new PageNumber());
                break;
            case "tab":
                EnsureRun();
                currentRun!.Append(new TabChar());
                break;
            case "uc":
                // Number of ANSI characters to skip after a following \uN control word
                if (cw.HasValue)
                {
                    try
                    {
                        runState.Uc = Math.Max(0, cw.Value!.Value);
                    }
                    catch
                    {
                        runState.Uc = 1;
                    }
                }
                else
                {
                    runState.Uc = 1;
                }
                break;
            case "u":
                if (cw.HasValue)
                {
                    int charCode = cw.Value!.Value;
                    if (charCode < 0)
                    {
                        // Unicode values greater than 32767 are expressed as negative numbers.
                        // For example, U+F020 would be \u-4064 in RTF: 
                        // sum 65536 to get 61472.
                        charCode += 65536;
                    }
                    string s = char.ConvertFromUtf32(charCode);
                    HandleText(s);
                    // After emitting the Unicode character, the RTF specification says that
                    // the following "uc" ANSI characters should be ignored. Track how many
                    // to skip on the formatting state so subsequent text tokens can consume them.
                    runState.PendingAnsiSkip = runState.Uc > 0 ? runState.Uc : 0;
                }
                break;
        }
        return false;
    }
}