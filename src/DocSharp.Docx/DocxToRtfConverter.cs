using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Collections;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public class DocxToRtfConverter : DocxConverterBase
{
    private FastStringCollection fonts = new FastStringCollection(); 
    private FastStringCollection colors = new FastStringCollection();

    internal override void ProcessBody(Body body, StringBuilder sb)
    {
        sb.Append(@"{\rtf1\ansi\deff0");
        sb.Append(@"{\fonttbl{\f0\fnil\fcharset0 Arial;}");        
        var bodySb = new StringBuilder();
        base.ProcessBody(body, bodySb);

        foreach (var font in fonts)
        {
            sb.Append(@"{\f" + font.Value + @"\fnil\fcharset0 " + font.Key + ";}");
        }
        sb.AppendLine("}");
        sb.Append(@"{\colortbl ;");
        foreach (var color in colors)
        {
            // Use black a last resort
            sb.Append(StringHelpers.ConvertToRtfColor(color.Key) ?? @"\red255\green255\blue255;");
        }
        sb.AppendLine("}");
        sb.Append(bodySb.ToString());
        sb.AppendLine("}");
    }

    internal override void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        //sb.Append(@"\pard ");
        base.ProcessParagraph(paragraph, sb);
        sb.Append(@"\par ");
        sb.AppendLine();
    }

    internal override void ProcessTable(Table table, StringBuilder sb)
    {
        //sb.AppendLine(@"{\trowd \trgaph108\trleft-108");

        //foreach (var row in table.Elements<TableRow>())
        //{
        //    foreach (var cell in row.Elements<TableCell>())
        //    {
        //        sb.Append(@"\cellx" + (cell.Descendants<Paragraph>().Count() * 1000)); // esempio di calcolo dimensione cella
        //        foreach (var paragraph in cell.Elements<Paragraph>())
        //        {
        //            ProcessParagraph(paragraph, sb);
        //        }
        //    }
        //    sb.AppendLine(@"\row");
        //}

        //sb.AppendLine(@"\pard \par}");
    }

    internal override void ProcessRun(Run run, StringBuilder sb)
    {
        var stylesPart = OpenXmlHelpers.GetMainDocumentPart(run)?.StyleDefinitionsPart?.Styles;
        var defaultRunStyle = stylesPart?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;

        var properties = run.GetFirstChild<RunProperties>();
        var text = run.GetFirstChild<Text>()?.InnerText;
        bool hasText = !string.IsNullOrEmpty(text);

        bool isBold, isItalic, isSingleStrike, isDoubleStrike, isSubscript, isSuperscript;
        isBold = isItalic = isSingleStrike = isDoubleStrike = isSubscript = isSuperscript = false;

        UnderlineValues? underlineValue = null;
        HighlightColorValues? highlightValue = null;

        bool isShadow, isOutline, isEmboss, isEngraveImprint;
        isShadow = isOutline = isEmboss = isEngraveImprint = false;

        bool isSmallCaps, isAllCaps, isHidden;
        isSmallCaps = isAllCaps = isHidden = false;

        if (hasText)
        {
            StyleRunProperties? runStyle = null;
            if (properties?.RunStyle?.Val?.Value is string styleId)
            {
                runStyle = stylesPart?.Elements<Style>().FirstOrDefault(s => s.StyleName?.Val == styleId)?.StyleRunProperties;               
            }

            isBold = runStyle?.Bold != null || 
                     defaultRunStyle?.Bold != null ||
                     properties?.Bold != null;
            isItalic = runStyle?.Italic != null ||
                       defaultRunStyle?.Italic != null ||
                       properties?.Italic != null;

            // Run properties takes precedence over style
            underlineValue = properties?.Underline?.Val ??
                             runStyle?.Underline?.Val ?? 
                             defaultRunStyle?.Underline?.Val ?? 
                             UnderlineValues.None;

            highlightValue = properties?.Highlight?.Val ?? HighlightColorValues.None;

            //underlineColor = properties?.Underline?.Color ??
            //                runStyle?.Underline?.Color ??
            //                defaultRunStyle?.Underline?.Color;

            isDoubleStrike =  runStyle?.DoubleStrike != null ||
                              defaultRunStyle?.DoubleStrike != null ||
                              properties?.DoubleStrike != null;
            isSingleStrike = !isDoubleStrike && (runStyle?.Strike != null ||
                                                defaultRunStyle?.Strike != null ||
                                                properties?.Strike != null);

            isSubscript = (runStyle?.VerticalTextAlignment?.Val != null && runStyle.VerticalTextAlignment.Val == "subscript") ||
                          (defaultRunStyle?.VerticalTextAlignment?.Val != null && defaultRunStyle.VerticalTextAlignment.Val == "subscript") ||
                          (properties?.VerticalTextAlignment?.Val != null && properties.VerticalTextAlignment.Val == "subscript");
            isSuperscript = (!isSubscript) && ((runStyle?.VerticalTextAlignment?.Val != null && runStyle.VerticalTextAlignment.Val == "superscript") ||
                                              (defaultRunStyle?.VerticalTextAlignment?.Val != null && defaultRunStyle.VerticalTextAlignment.Val == "superscript") ||
                                              (properties?.VerticalTextAlignment?.Val != null && properties.VerticalTextAlignment.Val == "superscript"));
            
            isSmallCaps = runStyle?.SmallCaps != null ||
                          defaultRunStyle?.SmallCaps != null ||
                          properties?.SmallCaps != null;
            isAllCaps = (!isSmallCaps) && (runStyle?.Caps != null ||
                                           defaultRunStyle?.Caps != null ||
                                           properties?.Caps != null);
            isEmboss = runStyle?.Emboss != null ||
                       defaultRunStyle?.Emboss != null ||
                       properties?.Emboss != null;
            isEngraveImprint = runStyle?.Imprint != null ||
                               defaultRunStyle?.Imprint != null ||
                               properties?.Imprint != null;
            isShadow = runStyle?.Shadow != null ||
                       defaultRunStyle?.Shadow != null ||
                       properties?.Shadow != null ||
                       properties?.Shadow14 != null;
            isOutline = runStyle?.Outline != null ||
                        defaultRunStyle?.Outline != null ||
                        properties?.Outline != null ||
                        properties?.TextOutlineEffect != null;
            isHidden = runStyle?.Vanish != null ||
                        defaultRunStyle?.Vanish != null ||
                        properties?.Vanish != null;

            // To be improved (Ascii value may not be present, although rare)
            string? font = properties?.RunFonts?.Ascii?.Value ??
                              runStyle?.RunFonts?.Ascii?.Value ??
                              defaultRunStyle?.RunFonts?.Ascii?.Value;
            if (!string.IsNullOrEmpty(font))
            {
                fonts.TryAddAndGetIndex(font, out int fontIndex);
                sb.Append($"\\f{fontIndex} ");
            }
            else
            {
                // Arial is already in the font table as last resort
                sb.Append(@"\f0 ");
            }

            string? color = properties?.Color?.Val ??
                            runStyle?.Color?.Val ??
                            defaultRunStyle?.Color?.Val;
            if (!string.IsNullOrEmpty(color))
            {
                colors.TryAddAndGetIndex(color, out int colorIndex);
                sb.Append($"\\cf{colorIndex} ");
            }
            else
            {
                // If no color is specified, \cf0 is automatically handled by word processors.
                // Note: for this reason the color table uses 1-based index,
                // while the font table should contain the f0 font.
                sb.Append(@"\cf0 ");
            }

            string? fontSize = properties?.FontSize?.Val ??
                            runStyle?.FontSize?.Val ??
                            defaultRunStyle?.FontSize?.Val;
            if (int.TryParse(fontSize, out int fs))
            {
                sb.Append($"\\fs{fs} ");
            }
            else
            {
                sb.Append(@"\fs22 ");
            }

            if (isItalic)
                sb.Append(@"\i ");

            if (isBold)
                sb.Append(@"\b ");

            if (isSingleStrike)
                sb.Append(@"\strike ");
            else if (isDoubleStrike)
                sb.Append(@"\striked1 ");

            if (underlineValue != null && underlineValue.HasValue && underlineValue.Value != UnderlineValues.None)
            {
                if (underlineValue.Value == UnderlineValues.Single)
                    sb.Append(@"\ul ");
                else if (underlineValue.Value == UnderlineValues.Dash)
                    sb.Append(@"\uldash ");
                else if (underlineValue.Value == UnderlineValues.Dotted)
                    sb.Append(@"\uld ");
                else if (underlineValue.Value == UnderlineValues.DotDash)
                    sb.Append(@"\uldashd ");
                else if (underlineValue.Value == UnderlineValues.DotDotDash)
                    sb.Append(@"\uldashdd ");
                else if (underlineValue.Value == UnderlineValues.DashLong)
                    sb.Append(@"\ulldash ");
                else if (underlineValue.Value == UnderlineValues.Double)
                    sb.Append(@"\uldb ");
                else if (underlineValue.Value == UnderlineValues.Thick)
                    sb.Append(@"\ulth ");
                else if (underlineValue.Value == UnderlineValues.DashedHeavy)
                    sb.Append(@"\ulthdash ");
                else if (underlineValue.Value == UnderlineValues.DottedHeavy)
                    sb.Append(@"\ulthd ");
                else if (underlineValue.Value == UnderlineValues.DashDotHeavy)
                    sb.Append(@"\ulthdashd ");
                else if (underlineValue.Value == UnderlineValues.DashDotDotHeavy)
                    sb.Append(@"\ulthdashdd ");
                else if (underlineValue.Value == UnderlineValues.DashLongHeavy)
                    sb.Append(@"\ulthldash ");
                else if (underlineValue.Value == UnderlineValues.Words)
                    sb.Append(@"\ulw ");
                else if (underlineValue.Value == UnderlineValues.Wave)
                    sb.Append(@"\ulwave ");
                else if (underlineValue.Value == UnderlineValues.WavyDouble)
                    sb.Append(@"\ululdbwave ");
                else if (underlineValue.Value == UnderlineValues.WavyHeavy)
                    sb.Append(@"\ulhwave ");
            }

            if (highlightValue != null && highlightValue.HasValue && highlightValue.Value != HighlightColorValues.None)
            {
                int highlightIndex = 0;
                if (highlightValue.Value == HighlightColorValues.Black)
                {
                    colors.TryAddAndGetIndex("000000", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.White)
                {
                    colors.TryAddAndGetIndex("FFFFFF", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Red)
                {
                    colors.TryAddAndGetIndex("FF0000", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Green)
                {
                    colors.TryAddAndGetIndex("00FF00", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Blue)
                {
                    colors.TryAddAndGetIndex("0000FF", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Yellow)
                {
                    colors.TryAddAndGetIndex("FFFF00", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Cyan)
                {
                    colors.TryAddAndGetIndex("00FFFF", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.Magenta)
                {
                    colors.TryAddAndGetIndex("FF00FF", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkRed)
                {
                    colors.TryAddAndGetIndex("800000", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkGreen)
                {
                    colors.TryAddAndGetIndex("008000", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkBlue)
                {
                    colors.TryAddAndGetIndex("000080", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkYellow)
                {
                    colors.TryAddAndGetIndex("808000", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkMagenta)
                {
                    colors.TryAddAndGetIndex("800080", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkCyan)
                {
                    colors.TryAddAndGetIndex("008080", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.DarkGray)
                {
                    colors.TryAddAndGetIndex("808080", out highlightIndex);
                }
                else if (highlightValue == HighlightColorValues.LightGray)
                {
                    colors.TryAddAndGetIndex("c0c0c0", out highlightIndex);
                }
                if (highlightIndex > 0)
                {
                    sb.Append($"\\highlight{highlightIndex}");
                }
            }

            if (isSubscript)
                sb.Append(@"\sub ");
            else if (isSuperscript)
                sb.Append(@"\super ");

            if (isEmboss)
                sb.Append(@"\embo ");

            if (isEngraveImprint)
                sb.Append(@"\impr ");

            if (isShadow)
                sb.Append(@"\shad ");

            if (isOutline)
                sb.Append(@"\outl ");

            if (isSmallCaps)
                sb.Append(@"\scaps ");
            else if (isAllCaps)
                sb.Append(@"\caps ");

            if (isHidden)
                sb.Append(@"\v ");
        }

        foreach (var element in run.Elements())
        {
            base.ProcessRunElement(element, sb);
        }

        if (hasText)
        {
            if (isHidden)
                sb.Append(@"\v0 ");

            if (isSmallCaps)
                sb.Append(@"\scaps0 ");
            else if (isAllCaps)
                sb.Append(@"\caps0 ");

            if (isOutline)
                sb.Append(@"\outl0 ");

            if (isShadow)
                sb.Append(@"\shad0 ");

            if (isEngraveImprint)
                sb.Append(@"\impr0 ");

            if (isEmboss)
                sb.Append(@"\embo0 ");

            if (isSubscript || isSuperscript)
                sb.Append(@"\nosupersub ");

            if (highlightValue != null && highlightValue.HasValue && highlightValue.Value != HighlightColorValues.None)
                sb.Append(@"\highlight0 ");

            if (underlineValue != null && underlineValue.HasValue && underlineValue.Value != UnderlineValues.None)
                sb.Append(@"\ul0 ");

            if (isSingleStrike)
                sb.Append(@"\strike0 ");
            else if (isDoubleStrike)
                sb.Append(@"\striked0 ");

            if (isBold)
                sb.Append(@"\b0 ");

            if (isItalic)
                sb.Append(@"\i0 ");
        }
    }

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        string escapedText = StringHelpers.ConvertToRtfUnicode(text.InnerText);
        sb.Append(escapedText);
    }

    internal override void ProcessPicture(Picture picture, StringBuilder sb)
    {
        //sb.Append(@"{\pict\wmetafile8\picwgoal100\pichgoal100 }");
    }

    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {
        //sb.Append(@"{\shp{\*\shpinst\shpleft0\shptop0\shpbottom0\shpright0 }}");
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {
        //var text = hyperlink.InnerText;
        //var uri = hyperlink.GetAttribute("anchor", "").Value;
        //sb.Append($@"{{\field{{\*\fldinst HYPERLINK ""{uri}"" }}{{\fldrslt {text}}}}}");
    }

    internal override void ProcessBookmark(BookmarkStart bookmark, StringBuilder sb)
    {
    }

    internal override void ProcessBreak(Break picture, StringBuilder sb)
    {
    }
}

