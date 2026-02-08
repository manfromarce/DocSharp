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
    private bool ProcessBreakControlWord(RtfControlWord cw, FormattingState runState)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch(name)
        {
            case "line":
                // text-wrapping line break. Avoid emitting duplicate breaks when previous token
                // already produced a text-wrapping break (some RTF producers emit both \line and \lbr).
                EnsureRun();
                if (!runState.LastWasLineBreak)
                {
                    currentRun!.Append(new Break() { Type = BreakValues.TextWrapping });
                    runState.LastWasLineBreak = true;
                }
                return true;
            case "page":
            case "column":
                // page/column breaks are distinct; reset the line-break flag.
                EnsureRun();
                currentRun!.Append(new Break() { Type = name == "page" ? BreakValues.Page : BreakValues.Column });
                runState.LastWasLineBreak = false;
                return true;
            case "lbr":
                // line break 
                if (cw.HasValue && !runState.LastWasLineBreak)
                {
                    if (cw.Value!.Value == 0)
                    {
                        EnsureRun();
                        currentRun!.Append(new Break() { Type = BreakValues.TextWrapping, Clear = BreakTextRestartLocationValues.None });
                        runState.LastWasLineBreak = true;
                    }
                    else if (cw.Value.Value == 1)
                    {
                        EnsureRun();
                        currentRun!.Append(new Break() { Type = BreakValues.TextWrapping, Clear = BreakTextRestartLocationValues.Left });
                        runState.LastWasLineBreak = true;
                    }
                    else if (cw.Value.Value == 2)
                    {
                        EnsureRun();
                        currentRun!.Append(new Break() { Type = BreakValues.TextWrapping, Clear = BreakTextRestartLocationValues.Right });
                        runState.LastWasLineBreak = true;
                    }
                    else if (cw.Value.Value == 3)
                    {
                        EnsureRun();
                        currentRun!.Append(new Break() { Type = BreakValues.TextWrapping, Clear = BreakTextRestartLocationValues.All });
                        runState.LastWasLineBreak = true;
                    }
                }
                return true;
        }
        return false;
    }
}