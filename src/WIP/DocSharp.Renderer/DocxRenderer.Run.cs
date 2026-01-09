using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocSharp.Docx;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using QuestPDF.Fluent;
using System.Globalization;
using M = DocumentFormat.OpenXml.Math;
using System.Diagnostics;

namespace DocSharp.Renderer;

public partial class DocxRenderer : DocxEnumerator<QuestPdfModel>, IDocumentRenderer<QuestPDF.Fluent.Document>
{
    internal override void ProcessRun(Run run, QuestPdfModel output)
    {
        if (run.GetEffectiveProperty<Vanish>().ToBool())
            return; // don't process hidden runs

        // Process run properties and add them to a new QuestPdfSpan object
        bool bold = run.GetEffectiveProperty<Bold>().ToBool();
        bool italic = run.GetEffectiveProperty<Italic>().ToBool();
        
        UnderlineStyle underline = UnderlineStyle.None;
        bool thickUnderline = false;
        QuestPDF.Infrastructure.Color? underlineColor = null;
        if (run.GetEffectiveProperty<Underline>() is Underline u && u.Val != null && u.Val.Value != UnderlineValues.None)
        {
            if (u.Val.Value == UnderlineValues.Dash || u.Val.Value == UnderlineValues.DashedHeavy || 
                u.Val.Value == UnderlineValues.DashLong || u.Val.Value == UnderlineValues.DashLongHeavy ||
                u.Val.Value == UnderlineValues.DotDash || u.Val.Value == UnderlineValues.DashDotDotHeavy ||
                u.Val.Value == UnderlineValues.DotDotDash || u.Val.Value == UnderlineValues.DashDotDotHeavy)
                underline = UnderlineStyle.Dashed;
            else if (u.Val.Value == UnderlineValues.Dotted || u.Val.Value == UnderlineValues.DottedHeavy)
                underline = UnderlineStyle.Dotted;
            else if (u.Val.Value == UnderlineValues.Wave || u.Val.Value == UnderlineValues.WavyDouble || u.Val.Value == UnderlineValues.WavyHeavy)
                underline = UnderlineStyle.Wavy;
            else if (u.Val.Value == UnderlineValues.Double)
                underline = UnderlineStyle.Double;
            else // solid, thick, words
                underline = UnderlineStyle.Solid;
            
            thickUnderline = u.Val.Value == UnderlineValues.DashedHeavy || u.Val.Value == UnderlineValues.DashLongHeavy || 
                             u.Val.Value == UnderlineValues.DashDotDotHeavy || u.Val.Value == UnderlineValues.DashDotDotHeavy || 
                             u.Val.Value == UnderlineValues.DottedHeavy || u.Val.Value == UnderlineValues.WavyHeavy || 
                             u.Val.Value == UnderlineValues.Thick;

            if (ColorHelpers.EnsureHexColor(u.Color?.Value) is string uc)
            {
                underlineColor = QuestPDF.Infrastructure.Color.FromHex(uc);
            }
        }
        StrikethroughStyle strikethrough = StrikethroughStyle.None;
        if (run.GetEffectiveProperty<Strike>().ToBool())
            strikethrough = StrikethroughStyle.Single;
        else if (run.GetEffectiveProperty<DoubleStrike>().ToBool()) 
            strikethrough = StrikethroughStyle.Double;

        SubSuperscript supSuperscript = SubSuperscript.Normal;
        var verticalPos = run.GetEffectiveProperty<VerticalTextAlignment>();
        if (verticalPos != null && verticalPos.Val != null && verticalPos.Val.Value != VerticalPositionValues.Baseline)
        {
            supSuperscript = verticalPos.Val.Value == VerticalPositionValues.Subscript ? SubSuperscript.Subscript : SubSuperscript.Superscript;
        }

        CapsType caps = CapsType.Normal;
        if (run.GetEffectiveProperty<SmallCaps>().ToBool())
            caps = CapsType.SmallCaps;
        else if (run.GetEffectiveProperty<Caps>().ToBool()) 
            caps = CapsType.AllCaps;

        float? fontSize = null;
        var fs = run.GetEffectiveProperty<FontSize>()?.Val?.Value;
        if (!string.IsNullOrEmpty(fs) && float.TryParse(fs, out float fontSizeValue))
        {
            fontSizeValue /= 2f; // Convert half-points to points
            fontSize = fontSizeValue;
        }

        // Text color
        QuestPDF.Infrastructure.Color? fontColor = null;
        var docxFontColor = run.GetEffectiveTextColor();
        if (!string.IsNullOrWhiteSpace(docxFontColor))
        {
            fontColor = QuestPDF.Infrastructure.Color.FromHex(docxFontColor!);
        }

        // Highlight and shading (highlight has priority over shading)
        QuestPDF.Infrastructure.Color? bgColor = null;
        var docxBgColor = run.GetEffectiveBackgroundColor();
        if (!string.IsNullOrWhiteSpace(docxBgColor))
        {
            bgColor = QuestPDF.Infrastructure.Color.FromHex(docxBgColor!);
        }

        string? fontFamily = null; 
        if (run.GetEffectiveProperty<RunFonts>()?.Ascii?.Value is string asciiFont && 
            !string.IsNullOrWhiteSpace(asciiFont))
        {
            fontFamily = asciiFont;
        }
        // TODO: improve fonts handling to support complex scripts;
        // check font embedding license; check QuestPDF subsetting options

        // TODO: letter spacing; vertical offset
        float? letterSpacing = null;

        var span = new QuestPdfSpan(null, bold, italic, underline, strikethrough, supSuperscript, caps, fontFamily, fontSize, fontColor, bgColor, underlineColor, letterSpacing, thickUnderline);

        // Add span to the paragraph/hyperlink.
        if (currentRunContainer.Count > 0)
            currentRunContainer.Peek().AddSpan(span);

        // Then, enumerate run elements (text, picture, break, page number, footnote reference...)
        currentSpan.Push(span);
        base.ProcessRun(run, output);
        if (currentSpan.Count > 0)
            currentSpan.Pop();
    }
}