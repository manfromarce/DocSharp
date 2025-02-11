using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Helpers;
using DocSharp.Docx.Rtf;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter
{
    private void ProcessTabs(Tabs tabs, StringBuilder sb)
    {
        foreach (var tab in tabs.Elements<TabStop>())
        {
            if (tab?.Val != null && tab.Val != TabStopValues.Clear &&
                tab?.Position != null && tab.Position.HasValue)
            {
                if (tab.Leader != null)
                {
                    if (tab.Leader.Value == TabStopLeaderCharValues.Dot)
                    {
                        sb.Append("\\tldot");
                    }
                    else if (tab.Leader.Value == TabStopLeaderCharValues.Heavy)
                    {
                        sb.Append("\\tlth");
                    }
                    else if (tab.Leader.Value == TabStopLeaderCharValues.Hyphen)
                    {
                        sb.Append("\\tlhyph");
                    }
                    else if (tab.Leader.Value == TabStopLeaderCharValues.MiddleDot)
                    {
                        sb.Append("\\tlmdot");
                    }
                    else if (tab.Leader.Value == TabStopLeaderCharValues.Underscore)
                    {
                        sb.Append("\\tlul");
                    }
                }
                if (tab.Val == TabStopValues.Bar)
                {
                    sb.Append($"\\tb{tab.Position.Value}");
                }
                else if (tab.Val == TabStopValues.Center)
                {
                    sb.Append($"\\tqc\\tx{tab.Position.Value}");
                }
                else if (tab.Val == TabStopValues.Decimal)
                {
                    sb.Append($"\\tqdec\\tx{tab.Position.Value}");
                }
                else if (tab.Val == TabStopValues.Left ||
                         tab.Val == TabStopValues.Start)
                {
                    sb.Append($"\\tx{tab.Position.Value}");
                }
                else if (tab.Val == TabStopValues.Number)
                {
                    sb.Append($"\\tx{tab.Position.Value}");
                }
                else if (tab.Val == TabStopValues.Right ||
                         tab.Val == TabStopValues.End)
                {
                    sb.Append($"\\tqr\\tx{tab.Position.Value}");
                }
            }
        }
    }

    internal override void ProcessPositionalTab(PositionalTab positionalTab, StringBuilder sb)
    {
        if (positionalTab.Alignment != null && positionalTab.RelativeTo != null)
        {
            if (positionalTab.Leader != null && positionalTab.Leader.Value != AbsolutePositionTabLeaderCharValues.None) 
            {
                if (positionalTab.Leader.Value == AbsolutePositionTabLeaderCharValues.Dot)
                {
                    sb.Append("{\\ptabldot");
                }
                else if (positionalTab.Leader.Value == AbsolutePositionTabLeaderCharValues.Hyphen)
                {
                    sb.Append("{\\ptablminus");
                }
                else if (positionalTab.Leader.Value == AbsolutePositionTabLeaderCharValues.MiddleDot)
                {
                    sb.Append("{\\ptablmdot");
                }
                else if (positionalTab.Leader.Value == AbsolutePositionTabLeaderCharValues.Underscore)
                {
                    sb.Append("{\\ptabluscore");
                }
            }
            else
            {
                sb.Append("{\\ptablnone");
            }
            sb.Append(' ');
            bool relativeToMargin = positionalTab.RelativeTo.Value == AbsolutePositionTabPositioningBaseValues.Margin;
            if (positionalTab.Alignment.Value == AbsolutePositionTabAlignmentValues.Left)
            {
                sb.Append(relativeToMargin ? "\\pmartabql" : "\\pindtabql");
            }
            else if (positionalTab.Alignment.Value == AbsolutePositionTabAlignmentValues.Center)
            {
                sb.Append(relativeToMargin ? "\\pmartabqc" : "\\pindtabqc");
            }
            else
            {
                sb.Append(relativeToMargin ? "\\pmartabqr" : "\\pindtabqr");
            }
            sb.Append('}');
        }
        sb.Append("{\\ptabldot \\pindtabqr}");
    }
}
