using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace DocSharp.Docx;

public static class FormattingHelpers
{
    internal static ConditionalFormattingFlags GetFlags(this ConditionalFormatStyle? cfs)
    {
        if (cfs == null) return ConditionalFormattingFlags.None;

        ConditionalFormattingFlags flags = ConditionalFormattingFlags.None;
        if (cfs.FirstRow?.Value == true) flags |= ConditionalFormattingFlags.FirstRow;
        if (cfs.LastRow?.Value == true) flags |= ConditionalFormattingFlags.LastRow;
        if (cfs.FirstColumn?.Value == true) flags |= ConditionalFormattingFlags.FirstColumn;
        if (cfs.LastColumn?.Value == true) flags |= ConditionalFormattingFlags.LastColumn;
        if (cfs.OddHorizontalBand?.Value == true) flags |= ConditionalFormattingFlags.OddRowBanding;
        if (cfs.EvenHorizontalBand?.Value == true) flags |= ConditionalFormattingFlags.EvenRowBanding;
        if (cfs.OddVerticalBand?.Value == true) flags |= ConditionalFormattingFlags.OddColumnBanding;
        if (cfs.EvenVerticalBand?.Value == true) flags |= ConditionalFormattingFlags.EvenColumnBanding;
        if (cfs.FirstRowFirstColumn?.Value == true) flags |= ConditionalFormattingFlags.NorthWestCell;
        if (cfs.FirstRowLastColumn?.Value == true) flags |= ConditionalFormattingFlags.NorthEastCell;
        if (cfs.LastRowFirstColumn?.Value == true) flags |= ConditionalFormattingFlags.SouthWestCell;
        if (cfs.LastRowLastColumn?.Value == true) flags |= ConditionalFormattingFlags.SouthEastCell;

        if (flags == ConditionalFormattingFlags.None && cfs.Val?.Value != null)
        {
            string val = cfs.Val.Value.PadLeft(12, '0');
            if (val.Length == 12)
            {
                // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.conditionalformatstyle.val?view=openxml-3.0.1#documentformat-openxml-wordprocessing-conditionalformatstyle-val
                if (val[0] == '1') flags |= ConditionalFormattingFlags.FirstRow;
                if (val[1] == '1') flags |= ConditionalFormattingFlags.LastRow;
                if (val[2] == '1') flags |= ConditionalFormattingFlags.FirstColumn;
                if (val[3] == '1') flags |= ConditionalFormattingFlags.LastColumn;
                if (val[4] == '1') flags |= ConditionalFormattingFlags.OddColumnBanding;
                if (val[5] == '1') flags |= ConditionalFormattingFlags.EvenColumnBanding;
                if (val[6] == '1') flags |= ConditionalFormattingFlags.OddRowBanding;
                if (val[7] == '1') flags |= ConditionalFormattingFlags.EvenRowBanding;
                if (val[9] == '1') flags |= ConditionalFormattingFlags.NorthEastCell;
                if (val[8] == '1') flags |= ConditionalFormattingFlags.NorthWestCell;
                if (val[11] == '1') flags |= ConditionalFormattingFlags.SouthEastCell;
                if (val[10] == '1') flags |= ConditionalFormattingFlags.SouthWestCell;
            }
        }
        return flags;
    }

    internal static ConditionalFormattingFlags GetCombinedConditionalFormattingFlags(this OpenXmlElement element)
    {
        ConditionalFormattingFlags flags = ConditionalFormattingFlags.None;

        Paragraph? paragraph = null;
        TableRow? row = null;
        TableCell? cell = null;

        if (element is Run)
            paragraph = element.GetFirstAncestor<Paragraph>();
        else if (element is Paragraph)
            paragraph = element as Paragraph;

        if (paragraph != null)
            flags |= paragraph.ParagraphProperties?.ConditionalFormatStyle.GetFlags() ?? ConditionalFormattingFlags.None;

        if (element is TableCell)
            cell = element as TableCell;
        else if (paragraph != null)
            cell = paragraph.GetFirstAncestor<TableCell>();

        if (cell != null)
            flags |= cell.TableCellProperties?.ConditionalFormatStyle.GetFlags() ?? ConditionalFormattingFlags.None;

        if (element is TableRow)
            row = element as TableRow;
        else if (cell != null)
            row = cell.GetFirstAncestor<TableRow>();
        else if (paragraph != null)
            row = paragraph.GetFirstAncestor<TableRow>();

        if (row != null)
            // TODO: can it contain multiple ConditionalFormatStyle ?
            flags |= row.TableRowProperties?.GetFirstChild<ConditionalFormatStyle>().GetFlags() ?? ConditionalFormattingFlags.None;

        return flags;
    }

    public static List<TableStyleProperties>? GetConditionalFormattingStyles(Style tableStyle, ConditionalFormattingFlags conditionalFormattingFlags)
    {
        if (conditionalFormattingFlags == ConditionalFormattingFlags.None)
            return null;

        var typesToCheck = new List<TableStyleOverrideValues>();
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.FirstRow))
            typesToCheck.Add(TableStyleOverrideValues.FirstRow);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.LastRow))
            typesToCheck.Add(TableStyleOverrideValues.LastRow);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.FirstColumn))
            typesToCheck.Add(TableStyleOverrideValues.FirstColumn);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.LastColumn))
            typesToCheck.Add(TableStyleOverrideValues.LastColumn);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.OddRowBanding))
            typesToCheck.Add(TableStyleOverrideValues.Band1Horizontal);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.EvenRowBanding))
            typesToCheck.Add(TableStyleOverrideValues.Band2Horizontal);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.OddColumnBanding))
            typesToCheck.Add(TableStyleOverrideValues.Band1Vertical);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.EvenColumnBanding))
            typesToCheck.Add(TableStyleOverrideValues.Band2Vertical);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.NorthWestCell))
            typesToCheck.Add(TableStyleOverrideValues.NorthWestCell);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.NorthEastCell))
            typesToCheck.Add(TableStyleOverrideValues.NorthEastCell);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.SouthWestCell))
            typesToCheck.Add(TableStyleOverrideValues.SouthWestCell);
        if (conditionalFormattingFlags.HasFlag(ConditionalFormattingFlags.SouthEastCell))
            typesToCheck.Add(TableStyleOverrideValues.SouthEastCell);

        return tableStyle.Elements<TableStyleProperties>()
            .Where(x => x.Type != null && typesToCheck.Contains(x.Type.Value))
            .ToList();
    }

    internal static T? GetConditionalFormattingProperty<T>(Styles? stylesPart, Paragraph p, ConditionalFormattingFlags conditionalFormattingFlags) where T : OpenXmlElement
    {
        if (stylesPart != null && conditionalFormattingFlags != ConditionalFormattingFlags.None)
        {
            if (p.GetFirstAncestor<Table>() is Table table &&
                table.GetFirstChild<TableProperties>() is TableProperties tPr &&
                tPr.TableStyle?.Val != null)
            {
                var tableStyle = stylesPart.GetStyleFromId(tPr.TableStyle.Val, StyleValues.Table);
                T? propertyValue = null;
                while (tableStyle != null)
                {
                    var conditionalStyles = GetConditionalFormattingStyles(tableStyle, conditionalFormattingFlags);

                    if (conditionalStyles != null)
                    {
                        foreach (var conditionalStyle in conditionalStyles)
                        {
                            propertyValue = conditionalStyle?.RunPropertiesBaseStyle?.GetFirstChild<T>() ??
                                            conditionalStyle?.StyleParagraphProperties?.GetFirstChild<T>();

                            if (propertyValue != null)
                            {
                                return propertyValue;
                            }
                        }
                    }

                    // Check styles from which the current style inherits
                    tableStyle = stylesPart.GetBaseStyle(tableStyle);
                }
            }
        }
        return null;
    }

    public static T? GetEffectiveProperty<T>(this OpenXmlElement element, Styles? styles = null) where T : OpenXmlElement
    {
        if (element is Paragraph p)
        {
            return GetEffectiveProperty<T>(p, styles);
        }
        else if (element is Run run)
        {
            return GetEffectiveProperty<T>(run, styles);
        }
        else if (element is Table table)
        {
            return GetEffectiveProperty<T>(table, styles);
        }
        else if (element is TableRow tableRow)
        {
            return GetEffectiveProperty<T>(tableRow, styles);
        }
        else if (element is TableCell tableCell)
        {
            return GetEffectiveProperty<T>(tableCell, styles);
        }
        else if (element is RunProperties runPr)
        {
            return GetEffectiveProperty<T>(runPr, styles);
        }
        else if (element is NumberingSymbolRunProperties runPr2)
        {
            return runPr2.GetFirstChild<T>();
        }
        else
        {
            return null;
        }
    }

    private static void MergeAttributes(OpenXmlElement? target, OpenXmlElement? source, List<string> attributeNames)
    {
        if (target != null && source != null)
        {
            var attrsTarget = target.GetAttributes();
            var attrsSource = source.GetAttributes();
            if (attrsTarget != null && attrsSource != null)
            {
                foreach (var attr in attrsSource)
                {
                    // Important: attributes are only set if not already present, to preserve the correct priority (e.g. paragraph properties -> style -> base style -> ...)
                    if (attributeNames.Contains(attr.LocalName, StringComparer.OrdinalIgnoreCase) &&
                        !string.IsNullOrEmpty(attr.Value) &&
                        attrsTarget.FirstOrDefault(a => a.LocalName.Equals(attr.LocalName, StringComparison.OrdinalIgnoreCase)).Value == null)
                        target.SetAttribute(attr);
                }
                //or
                //foreach (var attributeName in attributeNames)
                //{
                //    var attr = attrsSource.FirstOrDefault(a => a.LocalName.Equals(attributeName, StringComparison.OrdinalIgnoreCase));
                //    if (attr.Value != null && attrsTarget.FirstOrDefault(a => a.LocalName.Equals(attributeName, StringComparison.OrdinalIgnoreCase)).Value == null)
                //        target.SetAttribute(attr);
                //}
            }
        }
    }

    public static PreviousParagraphProperties? GetListParagraphProperties(this Paragraph paragraph)
    {
        if (paragraph.ParagraphProperties?.NumberingProperties is NumberingProperties numProperties &&
            numProperties?.NumberingLevelReference?.Val != null &&
            numProperties?.NumberingId?.Val != null &&
            paragraph.GetNumberingPart()?.NumberingDefinitionsPart?.Numbering is Numbering numbering)
        {
            if (numbering.Elements<NumberingInstance>().FirstOrDefault(x => x.NumberID != null &&
                                                                            x.NumberID.Value == numProperties.NumberingId.Val)
                        is NumberingInstance num)
            {
                Level? level = null;
                // If NumberingInstance has a LevelOverride, use it.
                level = num.FirstOrDefault<LevelOverride>(x => x.Level?.LevelIndex != null &&
                                                          x.Level.LevelIndex == numProperties.NumberingLevelReference.Val)?.Level;
                // Otherwise get level from AbstractNum
                if (num.AbstractNumId?.Val != null)
                {
                    level ??= numbering.FirstOrDefault<AbstractNum>(x => x.AbstractNumberId != null && x.AbstractNumberId.Value == num.AbstractNumId.Val)?
                                       .FirstOrDefault<Level>(x => x.LevelIndex != null && x.LevelIndex == numProperties.NumberingLevelReference.Val);
                }
                // Get paragraph properties for list level
                return level?.PreviousParagraphProperties;
            }
        }
        return null;
    }

    public static SpacingBetweenLines? GetEffectiveSpacing(this Paragraph paragraph, Styles? stylesPart = null)
    {
        var res = new SpacingBetweenLines();
        var attributes = new List<string> { "before", "beforeLines",
                                            "after", "afterLines",
                                            "beforeAutoSpacing", "afterAutoSpacing",
                                            "line", "lineRule" };

        MergeAttributes(res, paragraph.ParagraphProperties?.SpacingBetweenLines, attributes);

        // Note: for all properties except left/firstLine/hanging indent, it's not clear if we should also check numbering styles
        // MergeAttributes(res, paragraph.GetListParagraphProperties()?.SpacingBetweenLines, attributes);

        stylesPart ??= paragraph.GetStylesPart();

        // Check conditional formatting, if any
        var conditionalFormattingType = paragraph.GetCombinedConditionalFormattingFlags();
        MergeAttributes(res, GetConditionalFormattingProperty<SpacingBetweenLines>(stylesPart, paragraph, conditionalFormattingType), attributes);

        // Check paragraph style
        var paragraphStyle = stylesPart.GetStyleFromId(paragraph.ParagraphProperties?.ParagraphStyleId?.Val, StyleValues.Paragraph);
        while (paragraphStyle != null)
        {
            MergeAttributes(res, paragraphStyle?.StyleParagraphProperties?.SpacingBetweenLines, attributes);

            // Check styles from which the current style inherits
            paragraphStyle = stylesPart.GetBaseStyle(paragraphStyle);
        }

        // Check table paragraph style
        if (paragraph.GetFirstAncestor<Table>() is Table table &&
            table.GetFirstChild<TableProperties>() is TableProperties tableProperties)
        {
            var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
            while (tableStyle != null)
            {
                MergeAttributes(res, tableStyle?.StyleParagraphProperties?.SpacingBetweenLines, attributes);

                // Check styles from which the current style inherits
                tableStyle = stylesPart.GetBaseStyle(tableStyle);
            }
        }

        // Check default paragraph style for the current document
        MergeAttributes(res, stylesPart.GetDefaultParagraphStyle()?.SpacingBetweenLines, attributes);

        // Check normal paragraph style for the current document
        MergeAttributes(res, stylesPart.GetNormalParagraphStyle()?.SpacingBetweenLines, attributes);

        return res;
    }

    public static Indentation? GetEffectiveIndent(this Paragraph paragraph, Styles? stylesPart = null)
    {
        var res = new Indentation();
        var attributes = new List<string> { "left", "leftChars",
                                        "right", "rightChars",
                                        "start", "startChars",
                                        "end", "endChars",
                                        "firstLine", "firstLineChars",
                                        "hanging", "hangingChars" };

        MergeAttributes(res, paragraph.ParagraphProperties?.Indentation, attributes);

        // Check list style (numbering part), if any
        MergeAttributes(res, paragraph.GetListParagraphProperties()?.Indentation, attributes);

        stylesPart ??= paragraph.GetStylesPart();

        // Check conditional formatting, if any
        var conditionalFormattingType = paragraph.GetCombinedConditionalFormattingFlags();
        MergeAttributes(res,  GetConditionalFormattingProperty<Indentation>(stylesPart, paragraph, conditionalFormattingType), attributes);

        // Check paragraph style
        var paragraphStyle = stylesPart.GetStyleFromId(paragraph.ParagraphProperties?.ParagraphStyleId?.Val, StyleValues.Paragraph);
        while (paragraphStyle != null)
        {
            MergeAttributes(res, paragraphStyle?.StyleParagraphProperties?.Indentation, attributes);

            // Check styles from which the current style inherits
            paragraphStyle = stylesPart.GetBaseStyle(paragraphStyle);
        }

        // Check table paragraph style
        if (paragraph.GetFirstAncestor<Table>() is Table table &&
            table.GetFirstChild<TableProperties>() is TableProperties tableProperties)
        {
            var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
            while (tableStyle != null)
            {
                MergeAttributes(res, tableStyle?.StyleParagraphProperties?.Indentation, attributes);

                // Check styles from which the current style inherits
                tableStyle = stylesPart.GetBaseStyle(tableStyle);
            }
        }

        // Check default paragraph style for the current document
        MergeAttributes(res, stylesPart.GetDefaultParagraphStyle()?.Indentation, attributes);
        
        // Check normal paragraph style for the current document
        MergeAttributes(res, stylesPart.GetNormalParagraphStyle()?.Indentation, attributes);
        return res;
    }

    public static ParagraphBorders? GetPreviousParagraphBorders(this Paragraph paragraph, Styles? stylesPart = null)
    {
        var previousParagraph = paragraph.PreviousSibling() as Paragraph;
        if (previousParagraph != null)
        {
            return previousParagraph.GetEffectiveBorders(stylesPart);
        }
        return null;
    }

    public static ParagraphBorders? GetNextParagraphBorders(this Paragraph paragraph, Styles? stylesPart = null)
    {
        var nextParagraph = paragraph.NextSibling() as Paragraph;
        if (nextParagraph != null)
        {
            return nextParagraph.GetEffectiveBorders(stylesPart);
        }
        return null;
    }

    public static ParagraphBorders? GetEffectiveBorders(this Paragraph paragraph, Styles? stylesPart = null)
    {
        // Similar to GetEffectiveProperty, but continue recursing even if ParagraphBorders has been found, 
        // in order to merge null borders (top/bottom/left/right/between/bar) from different levels of the style hierarchy.
        // If there are no null borders, stop at the first ParagraphBorders found to avoid unnecessary recursion.
        // Improves performance compared to calling GetEffectiveProperty<T> for each border type separately.

        if (paragraph == null) return null;

        // Check paragraph properties
        var borders = paragraph.ParagraphProperties?.GetFirstChild<ParagraphBorders>();
        if (borders != null && borders.TopBorder != null && borders.BottomBorder != null && borders.LeftBorder != null && borders.RightBorder != null && borders.BetweenBorder != null && borders.BarBorder != null)
        {
            return borders;
        }

        // Note: for all properties except left/firstLine/hanging indent, it's not clear if we should also check numbering styles.
        // Skip for now.
        //propertyValue = paragraph.GetListParagraphProperties()?.GetFirstChild<ParagraphBorders>();

        stylesPart ??= paragraph.GetStylesPart();
        
         // Check conditional formatting, if any
        var conditionalFormattingType = paragraph.GetCombinedConditionalFormattingFlags();
        var conditionalBorders = GetConditionalFormattingProperty<ParagraphBorders>(stylesPart, paragraph, conditionalFormattingType);
        if (conditionalBorders != null)
        {
            borders ??= new ParagraphBorders();
            borders.TopBorder ??= conditionalBorders.TopBorder;
            borders.BottomBorder ??= conditionalBorders.BottomBorder;
            borders.LeftBorder ??= conditionalBorders.LeftBorder;
            borders.RightBorder ??= conditionalBorders.RightBorder;
            borders.BetweenBorder ??= conditionalBorders.BetweenBorder;
            borders.BarBorder ??= conditionalBorders.BarBorder;
        }
        if (borders != null && borders.TopBorder != null && borders.BottomBorder != null && borders.LeftBorder != null && borders.RightBorder != null && borders.BetweenBorder != null && borders.BarBorder != null)
        {
            return borders;
        }

        // Check paragraph style
        var paragraphStyle = stylesPart.GetStyleFromId(paragraph.ParagraphProperties?.ParagraphStyleId?.Val, StyleValues.Paragraph);
        while (paragraphStyle != null)
        {
            var styleBorders = paragraphStyle.StyleParagraphProperties?.ParagraphBorders;
            if (styleBorders != null)
            {
                borders ??= new ParagraphBorders();
                borders.TopBorder ??= styleBorders.TopBorder;
                borders.BottomBorder ??= styleBorders.BottomBorder;
                borders.LeftBorder ??= styleBorders.LeftBorder;
                borders.RightBorder ??= styleBorders.RightBorder;
                borders.BetweenBorder ??= styleBorders.BetweenBorder;
                borders.BarBorder ??= styleBorders.BarBorder;
            }
            if (borders != null && borders.TopBorder != null && borders.BottomBorder != null && borders.LeftBorder != null && borders.RightBorder != null && borders.BetweenBorder != null && borders.BarBorder != null)
            {
                return borders;
            }

            // Check styles from which the current style inherits
            paragraphStyle = stylesPart.GetBaseStyle(paragraphStyle);
        }

        // Check table paragraph style
        if (paragraph.GetFirstAncestor<Table>() is Table table &&
            table.GetFirstChild<TableProperties>() is TableProperties tableProperties)
        {
            var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
            while (tableStyle != null)
            {
                var styleBorders = tableStyle.StyleParagraphProperties?.ParagraphBorders;
                if (styleBorders != null)
                {
                    borders ??= new ParagraphBorders();
                    borders.TopBorder ??= styleBorders.TopBorder;
                    borders.BottomBorder ??= styleBorders.BottomBorder;
                    borders.LeftBorder ??= styleBorders.LeftBorder;
                    borders.RightBorder ??= styleBorders.RightBorder;
                    borders.BetweenBorder ??= styleBorders.BetweenBorder;
                    borders.BarBorder ??= styleBorders.BarBorder;                    
                }

                if (borders != null && borders.TopBorder != null && borders.BottomBorder != null && borders.LeftBorder != null && borders.RightBorder != null && borders.BetweenBorder != null && borders.BarBorder != null)
                {
                    return borders;
                }

                // Check styles from which the current style inherits
                tableStyle = stylesPart.GetBaseStyle(tableStyle);
            }
        }

        // Check default paragraph style for the current document
        var defaultBorders = stylesPart.GetDefaultParagraphStyle()?.ParagraphBorders;
        if (defaultBorders != null)
        {
            borders ??= new ParagraphBorders();
            borders.TopBorder ??= defaultBorders.TopBorder;
            borders.BottomBorder ??= defaultBorders.BottomBorder;
            borders.LeftBorder ??= defaultBorders.LeftBorder;
            borders.RightBorder ??= defaultBorders.RightBorder;
            borders.BetweenBorder ??= defaultBorders.BetweenBorder;
            borders.BarBorder ??= defaultBorders.BarBorder;
        }

        // Check normal paragraph style for the current document
        defaultBorders = stylesPart.GetNormalParagraphStyle()?.ParagraphBorders;
        if (defaultBorders != null)
        {
            borders ??= new ParagraphBorders();
            borders.TopBorder ??= defaultBorders.TopBorder;
            borders.BottomBorder ??= defaultBorders.BottomBorder;
            borders.LeftBorder ??= defaultBorders.LeftBorder;
            borders.RightBorder ??= defaultBorders.RightBorder;
            borders.BetweenBorder ??= defaultBorders.BetweenBorder;
            borders.BarBorder ??= defaultBorders.BarBorder;
        }

        // Return borders in any case now. If some border types are still null, it means they are not defined at any level and should be treated as such in converters.
        return borders;
    }

    public static T? GetEffectiveBorder<T>(this Paragraph paragraph, Styles? stylesPart = null) where T : BorderType
    {
        // Check paragraph properties
        T? propertyValue = paragraph.ParagraphProperties?.ParagraphBorders?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Note: for all properties except left/firstLine/hanging indent, it's not clear if we should also check numbering styles
        //propertyValue = paragraph.GetListParagraphProperties()?.ParagraphBorders?.GetFirstChild<T>();
        //if (propertyValue != null)
        //{
        //    return propertyValue;
        //}

        stylesPart ??= paragraph.GetStylesPart();

        // Check conditional formatting, if any
        var conditionalFormattingType = paragraph.GetCombinedConditionalFormattingFlags();
        propertyValue = GetConditionalFormattingProperty<ParagraphBorders>(stylesPart, paragraph, conditionalFormattingType)?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Check paragraph style
        var paragraphStyle = stylesPart.GetStyleFromId(paragraph.ParagraphProperties?.ParagraphStyleId?.Val, StyleValues.Paragraph);
        while (paragraphStyle != null)
        {
            propertyValue = paragraphStyle.StyleParagraphProperties?.ParagraphBorders?.GetFirstChild<T>();
            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            paragraphStyle = stylesPart.GetBaseStyle(paragraphStyle);
        }

        // Check table paragraph style
        if (paragraph.GetFirstAncestor<Table>() is Table table &&
            table.GetFirstChild<TableProperties>() is TableProperties tableProperties)
        {
            var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
            while (tableStyle != null)
            {
                propertyValue = tableStyle.StyleParagraphProperties?.ParagraphBorders?.GetFirstChild<T>();
                if (propertyValue != null)
                {
                    return propertyValue;
                }

                // Check styles from which the current style inherits
                tableStyle = stylesPart.GetBaseStyle(tableStyle);
            }
        }

        // Check default paragraph style for the current document
        return stylesPart.GetDefaultParagraphStyle()?.ParagraphBorders?.GetFirstChild<T>() ?? 
               stylesPart.GetNormalParagraphStyle()?.ParagraphBorders?.GetFirstChild<T>();
    }

    public static bool BordersAreEqual(ParagraphBorders? borders1, ParagraphBorders? borders2)
    {
        if (borders1 == null && borders2 == null) return true;
        if (borders1 == null || borders2 == null) return false;

        return BordersAreEqual(borders1.LeftBorder, borders2.LeftBorder) &&
               BordersAreEqual(borders1.RightBorder, borders2.RightBorder) &&
               BordersAreEqual(borders1.TopBorder, borders2.TopBorder) &&
               BordersAreEqual(borders1.BottomBorder, borders2.BottomBorder) &&
               BordersAreEqual(borders1.BetweenBorder, borders2.BetweenBorder) &&
               BordersAreEqual(borders1.BarBorder, borders2.BarBorder);
    }

    public static bool BordersAreEqual(BorderType? border1, BorderType? border2)
    {
        if (border1 == null && border2 == null) return true;
        if (border1 == null || border2 == null) return false;

        // var styleEqual = border1.Val == border2.Val || (border1.Val == null && border2.Val == null);
        // var colorEqual = border1.Color == border2.Color || (border1.Color == null && border2.Color == null);
        // var sizeEqual = border1.Size == border2.Size || (border1.Size == null && border2.Size == null);
        // var spaceEqual = border1.Space == border2.Space || (border1.Space == null && border2.Space == null);
        // var themeColorEqual = border1.ThemeColor == border2.ThemeColor || (border1.ThemeColor == null && border2.ThemeColor == null);
        // var themeShadeEqual = border1.ThemeShade == border2.ThemeShade || (border1.ThemeShade == null && border2.ThemeShade == null);
        // var themeTintEqual = border1.ThemeTint == border2.ThemeTint || (border1.ThemeTint == null && border2.ThemeTint == null);
        // return styleEqual && colorEqual && sizeEqual && spaceEqual && themeColorEqual && themeShadeEqual && themeTintEqual;

        return ((border1.Val == null && border2.Val == null) || (border1.Val?.InnerText == border2.Val?.InnerText)) &&
               (border1.Color == border2.Color || (border1.Color == null && border2.Color == null)) &&
               (border1.Size == border2.Size || (border1.Size == null && border2.Size == null)) &&
               (border1.Space == border2.Space || (border1.Space == null && border2.Space == null)) &&
               (border1.ThemeColor == border2.ThemeColor || (border1.ThemeColor == null && border2.ThemeColor == null)) &&
               (border1.ThemeShade == border2.ThemeShade || (border1.ThemeShade == null && border2.ThemeShade == null)) &&
               (border1.ThemeTint == border2.ThemeTint || (border1.ThemeTint == null && border2.ThemeTint == null));
    }

    // Helper function to get paragraph formatting from paragraph properties, style or default style.
    public static T? GetEffectiveProperty<T>(this Paragraph paragraph, Styles? stylesPart = null) where T : OpenXmlElement
    {
        // Check paragraph properties
        T? propertyValue = paragraph.ParagraphProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Note: for all properties except left/firstLine/hanging indent, it's not clear if we should also check numbering styles
        //propertyValue = paragraph.GetListParagraphProperties()?.GetFirstChild<T>();
        //if (propertyValue != null)
        //{
        //    return propertyValue;
        //}

        stylesPart ??= paragraph.GetStylesPart();

        // Check conditional formatting, if any
        var conditionalFormattingType = paragraph.GetCombinedConditionalFormattingFlags();
        propertyValue = GetConditionalFormattingProperty<T>(stylesPart, paragraph, conditionalFormattingType);
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Check paragraph style
        var paragraphStyle = stylesPart.GetStyleFromId(paragraph.ParagraphProperties?.ParagraphStyleId?.Val, StyleValues.Paragraph);
        while (paragraphStyle != null)
        {
            propertyValue = paragraphStyle.StyleParagraphProperties?.GetFirstChild<T>();
            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            paragraphStyle = stylesPart.GetBaseStyle(paragraphStyle);
        }

        // Check table paragraph style
        if (paragraph.GetFirstAncestor<Table>() is Table table &&
            table.GetFirstChild<TableProperties>() is TableProperties tableProperties)
        {
            var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
            while (tableStyle != null)
            {
                propertyValue = tableStyle.StyleParagraphProperties?.GetFirstChild<T>();
                if (propertyValue != null)
                {
                    return propertyValue;
                }

                // Check styles from which the current style inherits
                tableStyle = stylesPart.GetBaseStyle(tableStyle);
            }
        }

        // Check default paragraph style for the current document
        return stylesPart.GetDefaultParagraphStyle()?.GetFirstChild<T>() ?? 
               stylesPart.GetNormalParagraphStyle()?.GetFirstChild<T>();
    }

    // Helper function to get run formatting from run/paragraph properties, style or default style.
    public static T? GetEffectiveProperty<T>(this Run run, Styles? stylesPart = null) where T : OpenXmlElement
    {
        // Check run properties
        T? propertyValue = run.RunProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        stylesPart ??= run.GetStylesPart();

        // Check conditional formatting, if any
        var paragraph = run.GetFirstAncestor<Paragraph>();
        var conditionalFormattingType = paragraph is null ? ConditionalFormattingFlags.None : paragraph.GetCombinedConditionalFormattingFlags();
        if (paragraph != null && conditionalFormattingType != ConditionalFormattingFlags.None)
        {
            propertyValue = GetConditionalFormattingProperty<T>(stylesPart, paragraph, conditionalFormattingType);
            if (propertyValue != null)
            {
                return propertyValue;
            }
        }

        // Check run style
        var runStyle = stylesPart.GetStyleFromId(run.RunProperties?.RunStyle?.Val, StyleValues.Character);
        while (runStyle != null)
        {
            propertyValue = runStyle.StyleRunProperties?.GetFirstChild<T>();
            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            runStyle = stylesPart.GetBaseStyle(runStyle);
        }

        // Check paragraph style
        var paragraphProperties = paragraph?.ParagraphProperties;
        var paragraphStyle = stylesPart.GetStyleFromId(paragraphProperties?.ParagraphStyleId?.Val, StyleValues.Paragraph);
        while (paragraphStyle != null)
        {
            // Check paragraph style run properties
            propertyValue = paragraphStyle.StyleRunProperties?.GetFirstChild<T>();
            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check linked style, if any
            var linkedStyleId = paragraphStyle.LinkedStyle?.Val;
            if (linkedStyleId != null)
            {
                var linkedStyle = stylesPart.GetStyleFromId(linkedStyleId, StyleValues.Character);
                if (linkedStyle != null)
                {
                    propertyValue = linkedStyle.StyleRunProperties?.GetFirstChild<T>();
                    if (propertyValue != null)
                    {
                        return propertyValue;
                    }
                }
            }

            // Check styles from which the current style inherits
            paragraphStyle = stylesPart.GetBaseStyle(paragraphStyle);
        }

        // Check table run style
        if (run.GetFirstAncestor<Table>() is Table table &&
            table.GetFirstChild<TableProperties>() is TableProperties tableProperties)
        {
            var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
            while (tableStyle != null)
            {
                propertyValue = tableStyle.StyleRunProperties?.GetFirstChild<T>();
                if (propertyValue != null)
                {
                    return propertyValue;
                }

                // Check styles from which the current style inherits
                tableStyle = stylesPart.GetBaseStyle(tableStyle);
            }
        }

        // Check default run style for the current document
        return stylesPart.GetDefaultRunStyle()?.GetFirstChild<T>() ?? stylesPart.GetNormalRunStyle()?.GetFirstChild<T>();
    }

    public static string? GetEffectiveFont(this Run run, Styles? stylesPart = null, MainDocumentPart? mainPart = null)
    {
        // Check run properties
        var propertyValue = run.RunProperties?.RunFonts?.Ascii;
        if (propertyValue != null)
        {
            return propertyValue;
        }

        if (run.RunProperties?.RunFonts?.AsciiTheme != null)
        {
            if (run.RunProperties.RunFonts.AsciiTheme.Value == ThemeFontValues.MajorAscii || 
                run.RunProperties.RunFonts.AsciiTheme.Value == ThemeFontValues.MajorHighAnsi)
            {
                mainPart ??= run.GetMainDocumentPart();
                propertyValue = mainPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MajorFont?.LatinFont?.Typeface?.Value;
                if (!string.IsNullOrEmpty(propertyValue))
                {
                    return propertyValue;
                }
            }
            else if (run.RunProperties.RunFonts.AsciiTheme.Value == ThemeFontValues.MinorAscii || 
                     run.RunProperties.RunFonts.AsciiTheme.Value == ThemeFontValues.MinorHighAnsi)
            {
                mainPart ??= run.GetMainDocumentPart();
                propertyValue = mainPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
                if (!string.IsNullOrEmpty(propertyValue))
                {
                    return propertyValue;
                }
            }
        }

        stylesPart ??= run.GetStylesPart();

        // Check conditional formatting, if any
        var paragraph = run.GetFirstAncestor<Paragraph>();
        var conditionalFormattingType = paragraph is null ? ConditionalFormattingFlags.None : paragraph.GetCombinedConditionalFormattingFlags();
        if (paragraph != null && conditionalFormattingType != ConditionalFormattingFlags.None)
        {
            propertyValue = GetConditionalFormattingProperty<RunFonts>(stylesPart, paragraph, conditionalFormattingType)?.Ascii;
            if (!string.IsNullOrEmpty(propertyValue))
            {
                return propertyValue;
            }
        }

        // Check run style
        var runStyle = stylesPart.GetStyleFromId(run.RunProperties?.RunStyle?.Val, StyleValues.Character);
        while (runStyle != null)
        {
            propertyValue = runStyle.StyleRunProperties?.RunFonts?.Ascii;
            if (!string.IsNullOrEmpty(propertyValue))
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            runStyle = stylesPart.GetBaseStyle(runStyle);
        }

        // Check paragraph style
        var paragraphProperties = paragraph?.ParagraphProperties;
        var paragraphStyle = stylesPart.GetStyleFromId(paragraphProperties?.ParagraphStyleId?.Val, StyleValues.Paragraph);
        while (paragraphStyle != null)
        {
            // Check paragraph style run properties
            propertyValue = paragraphStyle.StyleRunProperties?.RunFonts?.Ascii;
            if (!string.IsNullOrEmpty(propertyValue))
            {
                return propertyValue;
            }

            // Check linked style, if any
            var linkedStyleId = paragraphStyle.LinkedStyle?.Val;
            if (linkedStyleId != null)
            {
                var linkedStyle = stylesPart.GetStyleFromId(linkedStyleId, StyleValues.Character);
                if (linkedStyle != null)
                {
                    propertyValue = linkedStyle.StyleRunProperties?.RunFonts?.Ascii;
                    if (!string.IsNullOrEmpty(propertyValue))
                    {
                        return propertyValue;
                    }
                }
            }

            // Check styles from which the current style inherits
            paragraphStyle = stylesPart.GetBaseStyle(paragraphStyle);
        }

        // Check table run style
        if (run.GetFirstAncestor<Table>() is Table table &&
            table.GetFirstChild<TableProperties>() is TableProperties tableProperties)
        {
            var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
            while (tableStyle != null)
            {
                propertyValue = tableStyle.StyleRunProperties?.RunFonts?.Ascii;
                if (!string.IsNullOrEmpty(propertyValue))
                {
                    return propertyValue;
                }

                // Check styles from which the current style inherits
                tableStyle = stylesPart.GetBaseStyle(tableStyle);
            }
        }

        // Check default run style for the current document
        var defaultStyle = stylesPart.GetDefaultRunStyle();
        propertyValue = defaultStyle?.RunFonts?.Ascii;
        if (!string.IsNullOrEmpty(propertyValue)) return propertyValue;

        var normalStyle = stylesPart.GetNormalRunStyle();
        propertyValue = normalStyle?.RunFonts?.Ascii;
        if (!string.IsNullOrEmpty(propertyValue)) return propertyValue;

        if (defaultStyle?.RunFonts?.AsciiTheme != null)
        {
            if (defaultStyle.RunFonts.AsciiTheme.Value == ThemeFontValues.MajorAscii || 
                defaultStyle.RunFonts.AsciiTheme.Value == ThemeFontValues.MajorHighAnsi)
            {
                mainPart ??= run.GetMainDocumentPart();
                propertyValue = mainPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MajorFont?.LatinFont?.Typeface?.Value;
                if (propertyValue != null)
                {
                    return propertyValue;
                }
            }
            else if (defaultStyle.RunFonts.AsciiTheme.Value == ThemeFontValues.MinorAscii || 
                     defaultStyle.RunFonts.AsciiTheme.Value == ThemeFontValues.MinorHighAnsi)
            {
                mainPart ??= run.GetMainDocumentPart();
                propertyValue = mainPart?.ThemePart?.Theme?.ThemeElements?.FontScheme?.MinorFont?.LatinFont?.Typeface?.Value;
                if (propertyValue != null)
                {
                    return propertyValue;
                }
            }
        }
        return propertyValue;
    }

    /// <summary>
    /// Helper function to get text background color from cell/table properties, style or default style.
    /// Patterns are ignored, returning the primary color only, unless only the secondary color is used.
    /// </summary>
    /// <param name="run"></param>
    /// <param name="stylesPart"></param>
    /// <returns></returns>
    public static string? GetEffectiveBackgroundColor(this TableCell tableCell, Styles? stylesPart = null)
    {
        return tableCell.GetEffectiveProperty<Shading>().ToHexColor();
    }

    /// <summary>
    /// Helper function to get text background color from paragraph properties, style or default style.
    /// Patterns are ignored, returning the primary color only, unless only the secondary color is used.
    /// </summary>
    /// <param name="run"></param>
    /// <param name="stylesPart"></param>
    /// <returns></returns>
    public static string? GetEffectiveBackgroundColor(this Paragraph paragraph, Styles? stylesPart = null)
    {
        return paragraph.GetEffectiveProperty<Shading>().ToHexColor();
    }

    /// <summary>
    /// Helper function to get text background color from run/paragraph properties, style or default style, 
    /// considering both highlight and shading 
    /// (in case of shading, patterns are ignored unless only the secondary color is used).
    /// </summary>
    /// <param name="run"></param>
    /// <param name="stylesPart"></param>
    /// <returns></returns>
    public static string? GetEffectiveBackgroundColor(this Run run, Styles? stylesPart = null)
    {
        // Highlight has priority over shading in Word.

        // Check run properties
        var propertyValue = run.RunProperties?.Highlight.ToHexColor() ?? 
                            run.RunProperties?.Shading?.ToHexColor();
        if (!string.IsNullOrWhiteSpace(propertyValue))
        {
            return propertyValue;
        }

        // Get styles if not passed to the function.
        stylesPart ??= run.GetStylesPart();

        // Check conditional formatting, if any
        var paragraph = run.GetFirstAncestor<Paragraph>();
        var conditionalFormattingType = paragraph is null ? ConditionalFormattingFlags.None : paragraph.GetCombinedConditionalFormattingFlags();
        if (paragraph != null && conditionalFormattingType != ConditionalFormattingFlags.None)
        {
            propertyValue = GetConditionalFormattingProperty<W.Highlight>(stylesPart, paragraph, conditionalFormattingType)?.ToHexColor() ?? 
                            GetConditionalFormattingProperty<W.Shading>(stylesPart, paragraph, conditionalFormattingType)?.ToHexColor();
            if (!string.IsNullOrWhiteSpace(propertyValue))
            {
                return propertyValue;
            }
        }

        // Check run style
        var runStyle = stylesPart.GetStyleFromId(run.RunProperties?.RunStyle?.Val, StyleValues.Character);
        while (runStyle != null)
        {
            propertyValue = runStyle.StyleRunProperties?.GetFirstChild<W.Highlight>()?.ToHexColor() ?? 
                            runStyle.StyleRunProperties?.Shading?.ToHexColor();
            if (!string.IsNullOrWhiteSpace(propertyValue))
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            runStyle = stylesPart.GetBaseStyle(runStyle);
        }

        // Check paragraph style
        var paragraphProperties = paragraph?.ParagraphProperties;
        var paragraphStyle = stylesPart.GetStyleFromId(paragraphProperties?.ParagraphStyleId?.Val, StyleValues.Paragraph);
        while (paragraphStyle != null)
        {
            // Check paragraph style run properties
            propertyValue = paragraphStyle.StyleRunProperties?.GetFirstChild<W.Highlight>()?.ToHexColor() ?? 
                            paragraphStyle.StyleRunProperties?.Shading?.ToHexColor();
            if (!string.IsNullOrWhiteSpace(propertyValue))
            {
                return propertyValue;
            }

            // Check linked style, if any
            var linkedStyleId = paragraphStyle.LinkedStyle?.Val;
            if (linkedStyleId != null)
            {
                var linkedStyle = stylesPart.GetStyleFromId(linkedStyleId, StyleValues.Character);
                if (linkedStyle != null)
                {
                    propertyValue = linkedStyle.StyleRunProperties?.GetFirstChild<W.Highlight>()?.ToHexColor() ?? 
                                    linkedStyle.StyleRunProperties?.Shading?.ToHexColor();
                    if (!string.IsNullOrWhiteSpace(propertyValue))
                    {
                        return propertyValue;
                    }
                }
            }

            // Check styles from which the current style inherits
            paragraphStyle = stylesPart.GetBaseStyle(paragraphStyle);
        }

        // Check table run style
        if (run.GetFirstAncestor<Table>() is Table table &&
            table.GetFirstChild<TableProperties>() is TableProperties tableProperties)
        {
            var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
            while (tableStyle != null)
            {
                propertyValue = tableStyle.StyleRunProperties?.GetFirstChild<W.Highlight>()?.ToHexColor() ?? 
                                tableStyle.StyleRunProperties?.Shading?.ToHexColor();
                if (!string.IsNullOrWhiteSpace(propertyValue))
                {
                    return propertyValue;
                }

                // Check styles from which the current style inherits
                tableStyle = stylesPart.GetBaseStyle(tableStyle);
            }
        }
        // Check default run style for the current document
        var defaultStyle = stylesPart.GetDefaultRunStyle();
        propertyValue = defaultStyle?.GetFirstChild<W.Highlight>()?.ToHexColor() ?? 
                        defaultStyle?.Shading?.ToHexColor();
        if (!string.IsNullOrWhiteSpace(propertyValue))
        {
            return propertyValue;
        }

        // Check normal style for the current document
        var normalStyle = stylesPart.GetNormalRunStyle();
        return normalStyle?.GetFirstChild<W.Highlight>()?.ToHexColor() ?? 
               normalStyle?.Shading?.ToHexColor();
    }

    /// <summary>
    /// Helper function to get text color from run/paragraph properties, style or default style, 
    /// considering both regular color and fill effect color.
    /// </summary>
    /// <param name="run"></param>
    /// <param name="stylesPart"></param>
    /// <returns></returns>
    public static string? GetEffectiveTextColor(this Run run, Styles? stylesPart = null)
    {
        // Fill effect has priority over color in Word.

        // Check run properties
        var propertyValue = run.RunProperties?.FillTextEffect.ToHexColor() ?? 
                            run.RunProperties?.Color?.ToHexColor();
        if (!string.IsNullOrWhiteSpace(propertyValue))
        {
            return propertyValue;
        }

        // Get styles if not passed to the function.
        stylesPart ??= run.GetStylesPart();

        // Check conditional formatting, if any
        var paragraph = run.GetFirstAncestor<Paragraph>();
        var conditionalFormattingType = paragraph is null ? ConditionalFormattingFlags.None : paragraph.GetCombinedConditionalFormattingFlags();
        if (paragraph != null && conditionalFormattingType != ConditionalFormattingFlags.None)
        {
            propertyValue = GetConditionalFormattingProperty<W14.FillTextEffect>(stylesPart, paragraph, conditionalFormattingType)?.ToHexColor() ?? 
                            GetConditionalFormattingProperty<W.Color>(stylesPart, paragraph, conditionalFormattingType)?.ToHexColor();
            if (!string.IsNullOrWhiteSpace(propertyValue))
            {
                return propertyValue;
            }
        }

        // Check run style
        var runStyle = stylesPart.GetStyleFromId(run.RunProperties?.RunStyle?.Val, StyleValues.Character);
        while (runStyle != null)
        {
            propertyValue = runStyle.StyleRunProperties?.GetFirstChild<W14.FillTextEffect>()?.ToHexColor() ?? 
                            runStyle.StyleRunProperties?.Color?.ToHexColor();
            if (!string.IsNullOrWhiteSpace(propertyValue))
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            runStyle = stylesPart.GetBaseStyle(runStyle);
        }

        // Check paragraph style
        var paragraphProperties = paragraph?.ParagraphProperties;
        var paragraphStyle = stylesPart.GetStyleFromId(paragraphProperties?.ParagraphStyleId?.Val, StyleValues.Paragraph);
        while (paragraphStyle != null)
        {
            // Check paragraph style run properties
            propertyValue = paragraphStyle.StyleRunProperties?.GetFirstChild<W14.FillTextEffect>()?.ToHexColor() ?? 
                            paragraphStyle.StyleRunProperties?.Color?.ToHexColor();
            if (!string.IsNullOrWhiteSpace(propertyValue))
            {
                return propertyValue;
            }

            // Check linked style, if any
            var linkedStyleId = paragraphStyle.LinkedStyle?.Val;
            if (linkedStyleId != null)
            {
                var linkedStyle = stylesPart.GetStyleFromId(linkedStyleId, StyleValues.Character);
                if (linkedStyle != null)
                {
                    propertyValue = linkedStyle.StyleRunProperties?.GetFirstChild<W14.FillTextEffect>()?.ToHexColor() ?? 
                                    linkedStyle.StyleRunProperties?.Color?.ToHexColor();
                    if (!string.IsNullOrWhiteSpace(propertyValue))
                    {
                        return propertyValue;
                    }
                }
            }

            // Check styles from which the current style inherits
            paragraphStyle = stylesPart.GetBaseStyle(paragraphStyle);
        }

        // Check table run style
        if (run.GetFirstAncestor<Table>() is Table table &&
            table.GetFirstChild<TableProperties>() is TableProperties tableProperties)
        {
            var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
            while (tableStyle != null)
            {
                propertyValue = tableStyle.StyleRunProperties?.GetFirstChild<W14.FillTextEffect>()?.ToHexColor() ?? 
                                tableStyle.StyleRunProperties?.Color?.ToHexColor();
                if (!string.IsNullOrWhiteSpace(propertyValue))
                {
                    return propertyValue;
                }

                // Check styles from which the current style inherits
                tableStyle = stylesPart.GetBaseStyle(tableStyle);
            }
        }

        // Check default run style for the current document
        var defaultStyle = stylesPart.GetDefaultRunStyle();
        propertyValue = defaultStyle?.GetFirstChild<W14.FillTextEffect>()?.ToHexColor() ?? 
                        defaultStyle?.Color?.ToHexColor();
        if (!string.IsNullOrWhiteSpace(propertyValue))
        {
            return propertyValue;
        }

        // Check normal style for the current document
        var normalStyle = stylesPart.GetNormalRunStyle();
        return normalStyle?.GetFirstChild<W14.FillTextEffect>()?.ToHexColor() ?? 
               normalStyle?.Color?.ToHexColor();
    }

    public static T? GetEffectiveProperty<T>(this RunProperties runPr, Styles? stylesPart = null) where T : OpenXmlElement
    {
        // Check run properties
        T? propertyValue = runPr.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        stylesPart ??= runPr.GetStylesPart();

        // Check run style
        var runStyle = stylesPart.GetStyleFromId(runPr.RunStyle?.Val, StyleValues.Character) ??
                       stylesPart.GetStyleFromId(runPr.RunStyle?.Val, StyleValues.Paragraph);
        while (runStyle != null)
        {
            propertyValue = runStyle.StyleRunProperties?.GetFirstChild<T>();
            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            runStyle = stylesPart.GetBaseStyle(runStyle);
        }
        return null;
    }

    // Helper function to get table formatting from table properties or style.
    public static T? GetEffectiveProperty<T>(this Table table, Styles? stylesPart = null) where T : OpenXmlElement
    {
        // Check table properties
        var tableProperties = table.GetFirstChild<TableProperties>();
        T? propertyValue = tableProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        stylesPart ??= table.GetStylesPart();
        var conditionalFormattingFlags = table.GetCombinedConditionalFormattingFlags();

        // Check table style
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        while (tableStyle != null)
        {
            // Check conditional formatting, if any
            var conditionalStyles = GetConditionalFormattingStyles(tableStyle, conditionalFormattingFlags);
            if (conditionalStyles != null)
            {
                foreach (var conditionalStyle in conditionalStyles)
                {
                    propertyValue = conditionalStyle?.TableStyleConditionalFormattingTableProperties?.GetFirstChild<T>();

                    if (propertyValue != null)
                    {
                        return propertyValue;
                    }
                }
            }

            // Check regular table style
            propertyValue = tableStyle.StyleTableProperties?.GetFirstChild<T>();
            // (should we also check row properties here ?)

            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            tableStyle = stylesPart.GetBaseStyle(tableStyle);
        }

        return null;
    }

    // Helper function to get row formatting from row/table properties or style.
    public static T? GetEffectiveProperty<T>(this TableRow row, Styles? stylesPart = null) where T : OpenXmlElement
    {
        // Check standard properties
        T? propertyValue = row.TableRowProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Check exceptions to table properties
        var tablePropertiesExceptions = row.TablePropertyExceptions; // has exceptions to TableProperties
        propertyValue = tablePropertiesExceptions?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Check table properties (can contain properties such as TableJustification).
        var tableProperties = row.GetFirstAncestor<Table>()?.GetFirstChild<TableProperties>();
        propertyValue = tableProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        stylesPart ??= row.GetStylesPart();
        var conditionalFormattingFlags = row.GetCombinedConditionalFormattingFlags();

        // Check table style
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        while (tableStyle != null)
        {
            // Check conditional formatting, if any
            var conditionalStyles = GetConditionalFormattingStyles(tableStyle, conditionalFormattingFlags);
            if (conditionalStyles != null)
            {
                foreach (var conditionalStyle in conditionalStyles)
                {
                    propertyValue = conditionalStyle?.TableStyleConditionalFormattingTableRowProperties?.GetFirstChild<T>() ??
                                    conditionalStyle?.TableStyleConditionalFormattingTableProperties?.GetFirstChild<T>();

                    if (propertyValue != null)
                    {
                        return propertyValue;
                    }
                }
            }

            // Check regular table style
            propertyValue = tableStyle.TableStyleConditionalFormattingTableRowProperties?.GetFirstChild<T>() ??
                            tableStyle.StyleTableProperties?.GetFirstChild<T>();
            // (should we also check cell properties here ?)

            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            tableStyle = stylesPart.GetBaseStyle(tableStyle);
        }

        return null;
    }

    // Helper function to get cell formatting from cell/table properties or style.
    public static T? GetEffectiveProperty<T>(this TableCell cell, Styles? stylesPart = null) where T : OpenXmlElement
    {
        // Check cell properties
        T? propertyValue = cell.TableCellProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Check row properties
        var row = cell.GetFirstAncestor<TableRow>();
        var tableRowProperties = row?.TableRowProperties; // can have e.g. TableCellSpacing
        var tablePropertiesExceptions = row?.TablePropertyExceptions; // has exceptions to TableProperties and many of the same properties,
                                                                      // so it's considered less specific than row properties
        propertyValue = tableRowProperties?.GetFirstChild<T>() ??
                        tablePropertiesExceptions?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Check table properties (properties such as borders are of different type between cell and table,
        // but other properties like shading may be found).
        var tableProperties = cell.GetFirstAncestor<Table>()?.GetFirstChild<TableProperties>();
        propertyValue = tableProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        stylesPart ??= cell.GetStylesPart();
        var conditionalFormattingFlags = cell.GetCombinedConditionalFormattingFlags();

        // Check table style
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        while (tableStyle != null)
        {
            // Check conditional formatting, if any
            var conditionalStyles = GetConditionalFormattingStyles(tableStyle, conditionalFormattingFlags);
            if (conditionalStyles != null)
            {
                foreach (var conditionalStyle in conditionalStyles)
                {
                    propertyValue = conditionalStyle?.TableStyleConditionalFormattingTableCellProperties?.GetFirstChild<T>() ??
                                    conditionalStyle?.TableStyleConditionalFormattingTableRowProperties?.GetFirstChild<T>() ??
                                    conditionalStyle?.TableStyleConditionalFormattingTableProperties?.GetFirstChild<T>();

                    if (propertyValue != null)
                    {
                        return propertyValue;
                    }
                }
            }

            // Check regular table style
            propertyValue = tableStyle.StyleTableCellProperties?.GetFirstChild<T>() ??
                            tableStyle.TableStyleConditionalFormattingTableRowProperties?.GetFirstChild<T>() ??
                            tableStyle.StyleTableProperties?.GetFirstChild<T>();
            // (should TableStyleConditionalFormattingTableRowProperties have precedence over StyleTableCellProperties?)

            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            tableStyle = stylesPart.GetBaseStyle(tableStyle);
        }

        return null;
    }

    // Helper function to get a border (top, bottom, left, start, diagonal...) from cell/table/row properties or style.
    public static BorderType? GetEffectiveBorder(this TableCell cell, Primitives.BorderValue borderValue,
                                                 bool isRightToLeft = false, Styles? stylesPart = null)
    {
        //bool isFirstRow = rowNumber == 1;
        //bool isFirstColumn = columnNumber == 1;
        //bool isLastRow = rowNumber == rowCount;
        //bool isLastColumn = columnNumber == columnCount;

        var targetTypesCell = new List<Type>();
        var targetTypesTable = new List<Type>();
        switch (borderValue)
        {
            case Primitives.BorderValue.Left:
                targetTypesCell.Add(typeof(LeftBorder));
                targetTypesCell.Add(isRightToLeft ? typeof(EndBorder) : typeof(StartBorder));
                targetTypesTable.Add(typeof(InsideVerticalBorder));
                break;
            case Primitives.BorderValue.Start:
                targetTypesCell.Add(typeof(StartBorder));
                targetTypesTable.Add(typeof(InsideVerticalBorder));
                break;
            case Primitives.BorderValue.Right:
                targetTypesCell.Add(typeof(RightBorder));
                targetTypesCell.Add(isRightToLeft ? typeof(StartBorder) : typeof(EndBorder));
                targetTypesTable.Add(typeof(InsideVerticalBorder));
                break;
            case Primitives.BorderValue.End:
                targetTypesCell.Add(typeof(EndBorder));
                targetTypesTable.Add(typeof(InsideVerticalBorder));
                break;
            case Primitives.BorderValue.Top:
                targetTypesCell.Add(typeof(TopBorder));
                targetTypesTable.Add(typeof(InsideHorizontalBorder));
                break;
            case Primitives.BorderValue.Bottom:
                targetTypesCell.Add(typeof(BottomBorder));
                targetTypesTable.Add(typeof(InsideHorizontalBorder));
                break;
            case Primitives.BorderValue.TopLeftToBottomRightDiagonal:
                targetTypesCell.Add(typeof(TopLeftToBottomRightCellBorder));
                break;
            case Primitives.BorderValue.TopRightToBottomLeftDiagonal:
                targetTypesCell.Add(typeof(TopRightToBottomLeftCellBorder));
                break;
        }

        // The types should be checked in order to preserve the correct priority.
        // For example, left and right should have precedence over start and end as they are more specific,
        // same applies to left/right/bottom/top over inside borders.
        OpenXmlElement? res = null;
        foreach (var type in targetTypesCell)
        {
            res = cell.TableCellProperties?.TableCellBorders?.FirstOrDefault(element => element.GetType() == type);
            if (res != null)
            {
                return (BorderType)res;
            }
        }

        // Check row properties
        var row = cell.GetFirstAncestor<TableRow>();
        var tablePropertiesExceptions = row?.TablePropertyExceptions;
        foreach (var type in targetTypesTable)
        {
            res = tablePropertiesExceptions?.TableBorders?.FirstOrDefault(element => element.GetType() == type);
            if (res != null)
            {
                return (BorderType)res;
            }
        }

        // Check table properties
        var tableProperties = cell.GetFirstAncestor<Table>()?.GetFirstChild<TableProperties>();
        foreach (var type in targetTypesTable)
        {
            res = tableProperties?.TableBorders?.FirstOrDefault(element => element.GetType() == type);
            if (res != null)
            {
                return (BorderType)res;
            }
        }

        // Check table style
        stylesPart ??= cell.GetStylesPart();
        var conditionalFormattingFlags = cell.GetCombinedConditionalFormattingFlags();
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        while (tableStyle != null)
        {
            // Check conditional formatting, if any
            var conditionalStyles = GetConditionalFormattingStyles(tableStyle, conditionalFormattingFlags);
            if (conditionalStyles != null)
            {
                foreach (var conditionalStyle in conditionalStyles)
                {
                    res = conditionalStyle?.TableStyleConditionalFormattingTableCellProperties?.TableCellBorders?.FirstOrDefault(element => targetTypesCell.Contains(element.GetType())) ??
                          conditionalStyle?.TableStyleConditionalFormattingTableProperties?.TableBorders?.FirstOrDefault(element => targetTypesTable.Contains(element.GetType()));

                    if (res != null)
                    {
                        return (BorderType)res;
                    }
                }
            }

            // Check regular table style
            res = tableStyle.StyleTableProperties?.TableBorders?.FirstOrDefault(element => targetTypesTable.Contains(element.GetType()));
            // (StyleTableCellProperties and row properties do not contain borders according to Open XML spec)

            if (res != null)
            {
                return (BorderType)res;
            }

            // Check styles from which the current style inherits
            tableStyle = stylesPart.GetBaseStyle(tableStyle);
        }
        return null;
    }

    // Helper function to get a border (top, bottom, left, start, diagonal...) from table/row properties or style.
    public static T? GetEffectiveBorder<T>(this TableRow row, Styles? stylesPart = null) where T : BorderType
    {
        // Check row properties
        var tablePropertiesExceptions = row.TablePropertyExceptions;
        T? propertyValue = tablePropertiesExceptions?.TableBorders?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Check table properties
        var tableProperties = row.GetFirstAncestor<Table>()?.GetFirstChild<TableProperties>();
        propertyValue = tableProperties?.TableBorders?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        stylesPart ??= row.GetStylesPart();

        // Check table style
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        var conditionalFormattingFlags = row.GetCombinedConditionalFormattingFlags();
        while (tableStyle != null)
        {
            // Check conditional formatting, if any
            var conditionalStyles = GetConditionalFormattingStyles(tableStyle, conditionalFormattingFlags);
            if (conditionalStyles != null)
            {
                foreach (var conditionalStyle in conditionalStyles)
                {
                    propertyValue = conditionalStyle?.TableStyleConditionalFormattingTableProperties?.TableBorders?.GetFirstChild<T>();

                    if (propertyValue != null)
                    {
                        return propertyValue;
                    }
                }
            }

            // Check regular table style
            propertyValue = tableStyle.StyleTableProperties?.TableBorders?.GetFirstChild<T>();
            // (row properties do not contain borders according to Open XML spec)
            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            tableStyle = stylesPart.GetBaseStyle(tableStyle);
        }
        return null;
    }

    // Helper function to get a border (top, bottom, left, start, ...) from row properties or style.
    public static T? GetEffectiveBorder<T>(this Table table, Styles? stylesPart = null) where T : BorderType
    {
        // Check table properties
        var tableProperties = table.GetFirstChild<TableProperties>();
        T? propertyValue = tableProperties?.TableBorders?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        stylesPart ??= table.GetStylesPart();

        // Check table style
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        var conditionalFormattingFlags = table.GetCombinedConditionalFormattingFlags();
        while (tableStyle != null)
        {
            // Check conditional formatting, if any
            var conditionalStyles = GetConditionalFormattingStyles(tableStyle, conditionalFormattingFlags);
            if (conditionalStyles != null)
            {
                foreach (var conditionalStyle in conditionalStyles)
                {
                    propertyValue = conditionalStyle?.TableStyleConditionalFormattingTableProperties?.TableBorders?.GetFirstChild<T>();

                    if (propertyValue != null)
                    {
                        return propertyValue;
                    }
                }
            }

            // Check regular table style
            propertyValue = tableStyle.StyleTableProperties?.TableBorders?.GetFirstChild<T>();
            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            tableStyle = stylesPart.GetBaseStyle(tableStyle);
        }
        return null;
    }

    public static OpenXmlElement? GetEffectiveMargin(this TableCell cell, Primitives.MarginValue marginValue, bool isRightToLeft = false, Styles? stylesPart = null)
    {
        // This method does not check table default margins by design,
        // as that would create double margins in the cell. 
        // Instead, table default margins should be processed separately.
        var targetTypesCell = new List<Type>();
        switch (marginValue)
        {
            case Primitives.MarginValue.Left:
                targetTypesCell.Add(typeof(LeftMargin));
                targetTypesCell.Add(typeof(TableCellLeftMargin));
                targetTypesCell.Add(isRightToLeft ? typeof(EndMargin) : typeof(StartMargin));
                break;
            case Primitives.MarginValue.Right:
                targetTypesCell.Add(typeof(RightMargin));
                targetTypesCell.Add(typeof(TableCellRightMargin));
                targetTypesCell.Add(isRightToLeft ? typeof(StartMargin) : typeof(EndMargin));
                break;
            case Primitives.MarginValue.Start:
                targetTypesCell.Add(typeof(StartMargin));
                break;
            case Primitives.MarginValue.End:
                targetTypesCell.Add(typeof(EndMargin));
                break;
            case Primitives.MarginValue.Top:
                targetTypesCell.Add(typeof(TopMargin));
                break;
            case Primitives.MarginValue.Bottom:
                targetTypesCell.Add(typeof(BottomMargin));
                break;
        }

        // The types should be checked in order to preserve the correct priority.
        // For example, left and right should have precedence over start and end as they are more specific.
        OpenXmlElement? res = null;
        foreach (var type in targetTypesCell)
        {
            res = cell.TableCellProperties?.TableCellMargin?.FirstOrDefault(element => element.GetType() == type);
            if (res != null)
            {
                return res;
            }
        }

        // Check table style
        var tableProperties = cell.GetFirstAncestor<Table>()?.GetFirstChild<TableProperties>();
        stylesPart ??= cell.GetStylesPart();
        var conditionalFormattingFlags = cell.GetCombinedConditionalFormattingFlags();
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        while (tableStyle != null)
        {
            // Check conditional formatting, if any
            var conditionalStyles = GetConditionalFormattingStyles(tableStyle, conditionalFormattingFlags);
            if (conditionalStyles != null)
            {
                foreach (var conditionalStyle in conditionalStyles)
                {
                    res = conditionalStyle?.TableStyleConditionalFormattingTableCellProperties?.TableCellMargin?.FirstOrDefault(element => targetTypesCell.Contains(element.GetType()));

                    if (res != null)
                    {
                        return res;
                    }
                }
            }

            // Check regular table style
            res = tableStyle.StyleTableCellProperties?.TableCellMargin?.FirstOrDefault(element => targetTypesCell.Contains(element.GetType()));
            if (res != null)
            {
                return res;
            }

            // Check styles from which the current style inherits
            tableStyle = stylesPart.GetBaseStyle(tableStyle);
        }
        return null;
    }

    public static T? GetEffectiveMargin<T>(this TableRow row, Styles? stylesPart = null) where T : OpenXmlElement
    {
        // Check row properties
        var tablePropertiesExceptions = row.TablePropertyExceptions;
        T? propertyValue = tablePropertiesExceptions?.TableCellMarginDefault?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Check table properties
        var tableProperties = row.GetFirstAncestor<Table>()?.GetFirstChild<TableProperties>();
        propertyValue = tableProperties?.TableCellMarginDefault?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        stylesPart ??= row.GetStylesPart();

        // Check table style
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        var conditionalFormattingFlags = row.GetCombinedConditionalFormattingFlags();
        while (tableStyle != null)
        {
            // Check conditional formatting, if any
            var conditionalStyles = GetConditionalFormattingStyles(tableStyle, conditionalFormattingFlags);
            if (conditionalStyles != null)
            {
                foreach (var conditionalStyle in conditionalStyles)
                {
                    propertyValue = conditionalStyle?.TableStyleConditionalFormattingTableProperties?.TableCellMarginDefault?.GetFirstChild<T>();
                    if (propertyValue != null)
                    {
                        return propertyValue;
                    }
                }
            }

            // Check regular table style
            propertyValue = tableStyle.StyleTableProperties?.TableCellMarginDefault?.GetFirstChild<T>();
            // (row properties do not contain margins according to Open XML spec)
            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            tableStyle = stylesPart.GetBaseStyle(tableStyle);
        }
        return null;
    }
}
