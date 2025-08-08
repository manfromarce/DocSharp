using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Helpers;
using DocSharp.Docx.Rtf;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    private void ProcessTabs(Tabs tabs, RtfStringWriter sb)
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
                        sb.Write("\\tldot");
                    }
                    else if (tab.Leader.Value == TabStopLeaderCharValues.Heavy)
                    {
                        sb.Write("\\tlth");
                    }
                    else if (tab.Leader.Value == TabStopLeaderCharValues.Hyphen)
                    {
                        sb.Write("\\tlhyph");
                    }
                    else if (tab.Leader.Value == TabStopLeaderCharValues.MiddleDot)
                    {
                        sb.Write("\\tlmdot");
                    }
                    else if (tab.Leader.Value == TabStopLeaderCharValues.Underscore)
                    {
                        sb.Write("\\tlul");
                    }
                }
                if (tab.Val == TabStopValues.Bar)
                {
                    sb.Write($"\\tb{tab.Position.Value.ToStringInvariant()}");
                }
                else if (tab.Val == TabStopValues.Center)
                {
                    sb.Write($"\\tqc\\tx{tab.Position.Value.ToStringInvariant()}");
                }
                else if (tab.Val == TabStopValues.Decimal)
                {
                    sb.Write($"\\tqdec\\tx{tab.Position.Value.ToStringInvariant()}");
                }
                else if (tab.Val == TabStopValues.Left ||
                         tab.Val == TabStopValues.Start)
                {
                    sb.Write($"\\tx{tab.Position.Value.ToStringInvariant()}");
                }
                else if (tab.Val == TabStopValues.Number)
                {
                    sb.Write($"\\tx{tab.Position.Value.ToStringInvariant()}");
                }
                else if (tab.Val == TabStopValues.Right ||
                         tab.Val == TabStopValues.End)
                {
                    sb.Write($"\\tqr\\tx{tab.Position.Value.ToStringInvariant()}");
                }
            }
        }
    }

    internal override void ProcessPositionalTab(PositionalTab positionalTab, RtfStringWriter sb)
    {
        if (positionalTab.Alignment != null && positionalTab.RelativeTo != null)
        {
            if (positionalTab.Leader != null && positionalTab.Leader.Value != AbsolutePositionTabLeaderCharValues.None) 
            {
                if (positionalTab.Leader.Value == AbsolutePositionTabLeaderCharValues.Dot)
                {
                    sb.Write("{\\ptabldot");
                }
                else if (positionalTab.Leader.Value == AbsolutePositionTabLeaderCharValues.Hyphen)
                {
                    sb.Write("{\\ptablminus");
                }
                else if (positionalTab.Leader.Value == AbsolutePositionTabLeaderCharValues.MiddleDot)
                {
                    sb.Write("{\\ptablmdot");
                }
                else if (positionalTab.Leader.Value == AbsolutePositionTabLeaderCharValues.Underscore)
                {
                    sb.Write("{\\ptabluscore");
                }
            }
            else
            {
                sb.Write("{\\ptablnone");
            }
            sb.Write(' ');
            bool relativeToMargin = positionalTab.RelativeTo.Value == AbsolutePositionTabPositioningBaseValues.Margin;
            if (positionalTab.Alignment.Value == AbsolutePositionTabAlignmentValues.Left)
            {
                sb.Write(relativeToMargin ? "\\pmartabql" : "\\pindtabql");
            }
            else if (positionalTab.Alignment.Value == AbsolutePositionTabAlignmentValues.Center)
            {
                sb.Write(relativeToMargin ? "\\pmartabqc" : "\\pindtabqc");
            }
            else
            {
                sb.Write(relativeToMargin ? "\\pmartabqr" : "\\pindtabqr");
            }
            sb.Write('}');
        }
        sb.Write("{\\ptabldot \\pindtabqr}");
    }
}
