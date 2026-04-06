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
    private bool ProcessParagraphTab(RtfControlWord cw, ParagraphProperties targetProperties)
    {
        var tabs = targetProperties.Tabs; // Tabs can contain a list of TabStop
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
            case "tx":
                return true;
            case "tb":
                return true;

            case "tqr":
                return true;
            case "tqc":
                return true;
            case "tqdec":
                return true;

            case "tldot":
                return true;
            case "tlmdot":
                return true;
            case "tlhyph":
                return true;
            case "tlul":
                return true;
            case "tlth":
                return true;
            case "tleq":
                return true;
        }
        return false;
    }

    private void ProcessAbsoluteTab(RtfGroup group)
    {
        var tab = new PositionalTab();
        var firstCw = group.Tokens.OfType<RtfControlWord>().FirstOrDefault();
        var lastCw = group.Tokens.OfType<RtfControlWord>().LastOrDefault();

        if (lastCw != null)
        {
            if (lastCw.Equals("pmartabql"))
            {
                tab.RelativeTo = AbsolutePositionTabPositioningBaseValues.Margin;
                tab.Alignment = AbsolutePositionTabAlignmentValues.Left;
            }
            else if (lastCw.Equals("pmartabqc"))
            {
                tab.RelativeTo = AbsolutePositionTabPositioningBaseValues.Margin;
                tab.Alignment = AbsolutePositionTabAlignmentValues.Center;
            }
            else if (lastCw.Equals("pmartabqr"))
            {
                tab.RelativeTo = AbsolutePositionTabPositioningBaseValues.Margin;
                tab.Alignment = AbsolutePositionTabAlignmentValues.Right;
            }
            else if (lastCw.Equals("pindtabql"))
            {
                tab.RelativeTo = AbsolutePositionTabPositioningBaseValues.Indent;
                tab.Alignment = AbsolutePositionTabAlignmentValues.Left;
            }
            else if (lastCw.Equals("pindtabqc"))
            {
                tab.RelativeTo = AbsolutePositionTabPositioningBaseValues.Indent;
                tab.Alignment = AbsolutePositionTabAlignmentValues.Center;
            }
            else if (lastCw.Equals("pindtabqr"))
            {
                tab.RelativeTo = AbsolutePositionTabPositioningBaseValues.Indent;
                tab.Alignment = AbsolutePositionTabAlignmentValues.Right;
            }
        }

        if (firstCw != null)
        {
            if (firstCw.Equals("ptablnone"))
                tab.Leader = AbsolutePositionTabLeaderCharValues.None;
            else if (firstCw.Equals("ptabldot"))
                tab.Leader = AbsolutePositionTabLeaderCharValues.Dot;
            else if (firstCw.Equals("ptablminus"))
                tab.Leader = AbsolutePositionTabLeaderCharValues.Hyphen;
            else if (firstCw.Equals("ptabluscore"))
                tab.Leader = AbsolutePositionTabLeaderCharValues.Underscore;
            else if (firstCw.Equals("ptablmdot"))
                tab.Leader = AbsolutePositionTabLeaderCharValues.MiddleDot;
        }

        if (tab.HasChildren)
        {
            AddRun().Append(tab);
        }
    }
}