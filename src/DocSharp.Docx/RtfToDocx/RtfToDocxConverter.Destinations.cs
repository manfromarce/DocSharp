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
    private void ProcessPartContent(RtfGroup group, Func<OpenXmlCompositeElement> createPart)
    {
        // Save current state
        var oldParagraph = pendingParagraph;
        var oldFmtStack = fmtStack.Clone();

        // Set context to a different document part (header, footer, footnote, endnote)
        containers.Push(createPart());
        pendingParagraph = null;
        currentRun = null;
        fmtStack.Clear();

        // Add content to the specified part
        ConvertGroup(group);        
        
        // Restore previous context; create subsequent content in a new run
        containers.Pop();
        pendingParagraph = oldParagraph;
        fmtStack = oldFmtStack;
    }

    private void ParseFontTable(RtfDestination dest)
    {
        if (dest == null) return;
        foreach (var token in dest.Tokens)
        {
            if (token is RtfGroup entry)
            {
                int? idx = null;
                int? fcharset = null;
                int? cpg = null;
                var sb = new StringBuilder();
                foreach (var et in entry.Tokens)
                {
                    if (et is RtfControlWord ecw)
                    {
                        var nm = (ecw.Name ?? string.Empty).ToLowerInvariant();
                        if (nm == "f" && ecw.HasValue)
                        {
                            idx = ecw.Value;
                        }
                        else if (nm == "fcharset" && ecw.HasValue)
                        {
                            fcharset = ecw.Value;
                        }
                        else if (nm == "cpg" && ecw.HasValue)
                        {
                            cpg = ecw.Value;
                        }
                        continue;
                    }
                    if (et is RtfText etxt)
                    {
                        sb.Append(etxt.Text ?? string.Empty);
                    }
                }
                if (idx.HasValue)
                {
                    var name = sb.ToString().Trim();
                    // remove trailing semicolon used as delimiter in fonttbl entries
                    if (name.EndsWith(";")) name = name.Substring(0, name.Length - 1).Trim();
                    if (!string.IsNullOrEmpty(name))
                        fontTable[idx.Value] = new RtfFontInfo() { Name = name, FCharset = fcharset, CodePage = cpg };
                }
            }
        }
    }

    private void ParseColorTable(RtfDestination dest)
    {
        if (dest == null) return;

        int r = -1, g = -1, b = -1;
        foreach (var token in dest.Tokens)
        {
            if (token is RtfControlWord cw)
            {
                var n = (cw.Name ?? string.Empty).ToLowerInvariant();
                if (n == "red" && cw.HasValue) r = cw.Value ?? 0;
                else if (n == "green" && cw.HasValue) g = cw.Value ?? 0;
                else if (n == "blue" && cw.HasValue) b = cw.Value ?? 0;
            }
            else if (token is RtfText txt)
            {
                var s = txt.Text ?? string.Empty;
                foreach (var ch in s)
                {
                    if (ch == ';')
                    {
                        if (r == -1 && g == -1 && b == -1)
                        {
                            // If the first color entry is empty, it should be assumed as "auto" color 
                            // (usually black, we can avoid forcing black depending on the control word 
                            // when index 0 is referenced, but we add black here so that subsequent colors 
                            // are mapped to correct index).
                            colorTable.Add((0, 0, 0));
                        }
                        else
                        {
                            // End of color entry
                            colorTable.Add((Clamp(r), Clamp(g), Clamp(b)));
                            r = g = b = 0;                            
                        }
                    }
                }
            }
        }

        static int Clamp(int v) => Math.Max(0, Math.Min(255, v));
    }
}