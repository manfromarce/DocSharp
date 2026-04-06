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
    private TabStop? pendingTab;

    private bool ProcessParagraphTab(RtfControlWord cw, ParagraphProperties targetProperties)
    {
        var name = (cw.Name ?? string.Empty).ToLowerInvariant();
        switch (name)
        {
            case "tx":
                if (cw.HasValue)
                {
                    targetProperties.Tabs ??= new Tabs();
                    pendingTab ??= new TabStop();
                    pendingTab.Val ??= TabStopValues.Left;
                    pendingTab.Position = cw.Value!.Value;
                    targetProperties.Tabs.Append(pendingTab);
                    pendingTab = null;
                }
                return true;
            case "tb":
                if (cw.HasValue)
                {
                    targetProperties.Tabs ??= new Tabs();
                    pendingTab ??= new TabStop();
                    pendingTab.Val ??= TabStopValues.Bar;
                    pendingTab.Position = cw.Value!.Value;
                    targetProperties.Tabs.Append(pendingTab);
                    pendingTab = null;
                }
                return true;

            case "tqc":
                pendingTab ??= new TabStop();
                pendingTab.Val = TabStopValues.Center;
                return true;
            case "tqr":
                pendingTab ??= new TabStop();
                pendingTab.Val = TabStopValues.Right;
                return true;
            case "tqdec":
                pendingTab ??= new TabStop();
                pendingTab.Val = TabStopValues.Decimal;
                return true;

            case "tldot":
                pendingTab ??= new TabStop();
                pendingTab.Leader = TabStopLeaderCharValues.Dot;
                return true;
            case "tlmdot":
                pendingTab ??= new TabStop();
                pendingTab.Leader = TabStopLeaderCharValues.MiddleDot;
                return true;
            case "tlhyph":
                pendingTab ??= new TabStop();
                pendingTab.Leader = TabStopLeaderCharValues.Hyphen;
                return true;
            case "tlul":
                pendingTab ??= new TabStop();
                pendingTab.Leader = TabStopLeaderCharValues.Underscore;
                return true;
            case "tlth":
                pendingTab ??= new TabStop();
                pendingTab.Leader = TabStopLeaderCharValues.Heavy;
                return true;
            case "tleq":
                // Leader equal sign, not available in DOCX.
                // pendingTab ??= new TabStop();
                // pendingTab.Leader = TabStopLeaderCharValues.None;
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