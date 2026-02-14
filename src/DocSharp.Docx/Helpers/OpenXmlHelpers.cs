using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace DocSharp.Docx;

public static class OpenXmlHelpers
{
    public static void Clear(this OpenXmlElement element)
    {
        element.RemoveAllChildren();
        element.ClearAllAttributes();
    }

    public static void Clear<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        var subElements = element.Elements<T>().ToArray();
        int count = subElements.Length;
        for (int i = count - 1; i >= 0; i--)
        {
            var subElement = subElements[count - 1];
            subElement.Remove();
        }    
    }

    public static void ClearExcept<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        element.ClearAllAttributes();
        var subElements = element.Elements().ToArray();
        int count = subElements.Length;
        for (int i = count - 1; i >= 0; i--)
        {
            var subElement = subElements[count - 1];
            if (subElement is not T)
                subElement.Remove();
        }
    }

    public static int GetLinksCount(WordprocessingDocument document)
    {
        return document.MainDocumentPart?.HyperlinkRelationships.Count() ?? 0;
    }

    public static int GetImagesCount(WordprocessingDocument document)
    {
        return document.MainDocumentPart?.ImageParts.Count() ?? 0;
    }

    public static bool IsMathElement(this OpenXmlElement element)
    {
        return element.NamespaceUri.StartsWith(OpenXmlConstants.MathNamespace, StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsVmlElement(this OpenXmlElement element)
    {
        return element.NamespaceUri.StartsWith(OpenXmlConstants.VmlNamespace, StringComparison.OrdinalIgnoreCase);
    }

    public static T? NextElement<T>(this OpenXmlElement? element) where T : OpenXmlElement
    {
        if (element != null && element.GetFirstAncestor<Document>() is Document document)
        {
            return document.Descendants<T>().FirstOrDefault(x => x.IsAfter(element));
        }
        return null;
    }

    public static T? GetFirstAncestor<T>(this OpenXmlElement? element) where T : OpenXmlElement
    {
        if (element != null && element.Parent != null)
        {
            if (element.Parent is T result)
            {
                return result;
            }
            else
            {
                return GetFirstAncestor<T>(element.Parent);
            }
        }
        return null;
    }

    public static T? GetFirstDescendant<T>(this OpenXmlElement? element) where T : OpenXmlElement
    {
        return element?.Descendants<T>().FirstOrDefault();
    }

    /// <summary>
    /// Helper function to retrieve main document part from an Open XML element.
    /// </summary>
    /// <returns></returns>
    public static MainDocumentPart? GetMainDocumentPart(this OpenXmlElement element)
    {
        var root = element.GetRoot();
        if (root is OpenXmlPartRootElement rootElement)
        {
            return (rootElement.OpenXmlPart?.OpenXmlPackage as WordprocessingDocument)?.MainDocumentPart;
        }
        return null;
    }

    /// <summary>
    /// Helper function to retrieve main document part from an Open XML element.
    /// </summary>
    /// <returns></returns>
    public static OpenXmlPart? GetRootPart(this OpenXmlElement element)
    {
        var root = element.GetRoot();
        if (root is OpenXmlPartRootElement rootElement)
        {
            return rootElement.OpenXmlPart;
        }
        return null;
    }

    public static WordprocessingDocument? GetWordprocessingDocument(this OpenXmlElement element)
    {
        return element.GetRootPart()?.OpenXmlPackage as WordprocessingDocument;
    }

    public static OpenXmlElement GetRoot(this OpenXmlElement element)
    {
        OpenXmlElement rootElement = element;
        while (rootElement.Parent != null)
        {
            rootElement = rootElement.Parent;
        }
        return rootElement;
    }

    /// <summary>
    /// Helper function to retrieve styles part from an Open XML element.
    /// </summary>
    /// <returns></returns>
    public static Styles? GetStylesPart(this OpenXmlElement element)
    {
        return GetMainDocumentPart(element)?.StyleDefinitionsPart?.Styles;
    }

    /// <summary>
    /// Helper function to retrieve Numbering part from an Open XML element.
    /// </summary>
    /// <returns></returns>
    public static Numbering? GetNumberingPart(this OpenXmlElement element)
    {
        return GetMainDocumentPart(element)?.NumberingDefinitionsPart?.Numbering;
    }

    /// <summary>
    /// Helper function to retrieve Theme part from an Open XML element.
    /// </summary>
    /// <returns></returns>
    public static DocumentFormat.OpenXml.Drawing.Theme? GetThemePart(this OpenXmlElement element)
    {
        return GetMainDocumentPart(element)?.ThemePart?.Theme;
    }

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

    public static T? GetEffectiveProperty<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        if (element is Paragraph p)
        {
            return GetEffectiveProperty<T>(p);
        }
        else if (element is Run run)
        {
            return GetEffectiveProperty<T>(run);
        }
        else if (element is Table table)
        {
            return GetEffectiveProperty<T>(table);
        }
        else if (element is TableRow tableRow)
        {
            return GetEffectiveProperty<T>(tableRow);
        }
        else if (element is TableCell tableCell)
        {
            return GetEffectiveProperty<T>(tableCell);
        }
        else if (element is RunProperties runPr)
        {
            return GetEffectiveProperty<T>(runPr);
        }
        else if (element is NumberingSymbolRunProperties runPr2)
        {
            return GetEffectiveProperty<T>(runPr2);
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

        stylesPart ??= GetStylesPart(paragraph);

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
        return stylesPart.GetDefaultParagraphStyle()?.ParagraphBorders?.GetFirstChild<T>();
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
                level = num.Elements<LevelOverride>()
                    .FirstOrDefault(x => x.Level?.LevelIndex != null &&
                                            x.Level.LevelIndex == numProperties.NumberingLevelReference.Val)?.Level;
                // Otherwise get level from AbstractNum
                if (num.AbstractNumId?.Val != null)
                {
                    level ??= numbering.Elements<AbstractNum>().FirstOrDefault(x => x.AbstractNumberId != null &&
                                                                                x.AbstractNumberId.Value == num.AbstractNumId.Val)?
                                    .Elements<Level>()
                                    .FirstOrDefault(x => x.LevelIndex != null &&
                                                        x.LevelIndex == numProperties.NumberingLevelReference.Val);
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

        stylesPart ??= GetStylesPart(paragraph);

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

        stylesPart ??= GetStylesPart(paragraph);

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

        return res;
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

        stylesPart ??= GetStylesPart(paragraph);

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
        return stylesPart.GetDefaultParagraphStyle()?.GetFirstChild<T>();
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

        stylesPart ??= GetStylesPart(run);

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
        return stylesPart.GetDefaultRunStyle()?.GetFirstChild<T>();
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
        stylesPart ??= GetStylesPart(run);

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
        return defaultStyle?.GetFirstChild<W.Highlight>()?.ToHexColor() ?? 
               defaultStyle?.Shading?.ToHexColor();
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
        stylesPart ??= GetStylesPart(run);

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
        return defaultStyle?.GetFirstChild<W14.FillTextEffect>()?.ToHexColor() ?? 
               defaultStyle?.Color?.ToHexColor();
    }

    public static T? GetEffectiveProperty<T>(this RunProperties runPr, Styles? stylesPart = null) where T : OpenXmlElement
    {
        // Check run properties
        T? propertyValue = runPr.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        stylesPart ??= GetStylesPart(runPr);

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

    public static T? GetEffectiveProperty<T>(this NumberingSymbolRunProperties runPr) where T : OpenXmlElement
    {
        // Check run properties
        return runPr.GetFirstChild<T>();
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

        stylesPart ??= GetStylesPart(table);
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

        stylesPart ??= GetStylesPart(row);
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

        stylesPart ??= GetStylesPart(cell);
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
        stylesPart ??= GetStylesPart(cell);
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

        stylesPart ??= GetStylesPart(row);

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
        T? propertyValue = null;

        // Check table properties
        var tableProperties = table.GetFirstChild<TableProperties>();
        propertyValue = tableProperties?.TableBorders?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        stylesPart ??= GetStylesPart(table);

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
        stylesPart ??= GetStylesPart(cell);
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

        stylesPart ??= GetStylesPart(row);

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

    public static void InsertBeforeLastOfType<T>(this OpenXmlCompositeElement parent, OpenXmlElement element)
        where T : OpenXmlElement
    {
        var refElement = parent.Elements<T>().LastOrDefault();
        if (refElement == null)
        {
            parent.AppendChild(element);
        }
        else
        {
            parent.InsertBefore(element, refElement);
        }
    }

    public static void InsertAfterLastOfType<T>(this OpenXmlCompositeElement parent, OpenXmlElement element)
        where T : OpenXmlElement
    {
        var refElement = parent.Elements<T>().LastOrDefault();
        if (refElement == null)
        {
            parent.AppendChild(element);
        }
        else
        {
            parent.InsertAfter(element, refElement);
        }
    }

    /// <summary>
    /// Get page size of the last section in DXA (1/20th of point).
    /// </summary>
    /// <param name="document">The word processing document object</param>
    /// <returns>A Size in twenthies of a point.</returns>
    public static Size GetPageSize(this WordprocessingDocument document)
    {
        var sectionProperties = document.MainDocumentPart?.Document?.Body?.Elements<SectionProperties>().LastOrDefault();
        var pageSize = sectionProperties?.GetFirstChild<PageSize>();
        if (pageSize?.Width != null && pageSize?.Height != null)
        {
            uint width = pageSize.Width?.Value ?? 0;
            uint height = pageSize.Height?.Value ?? 0;
            if (pageSize.Orient != null && pageSize.Orient.Value == PageOrientationValues.Landscape)
            {
                return new Size((int)height, (int)width);
            }
            else
            {
                return new Size((int)width, (int)height);
            }
        }
        // If not found, return A4 portrait.
        // A4 page size is 210 mm x 297 mm = 11906 x 16838
        return new Size(11906, 16838);
    }

    /// <summary>
    /// Get page size of the last section in DXA (also called twips) (1/20th of point) excluding page margins.
    /// </summary>
    /// <param name="document">The word processing document object</param>
    /// <returns>A Size in twenthies of a point.</returns>
    public static Size GetEffectivePageSize(this WordprocessingDocument document)
    {
        var sectionProperties = document.MainDocumentPart?.Document?.Body?.Elements<SectionProperties>().LastOrDefault();
        var pageSize = sectionProperties?.GetFirstChild<PageSize>();
        var pageMargin = sectionProperties?.GetFirstChild<PageMargin>();

        if (pageSize?.Width != null && pageSize.Height != null)
        {
            uint leftMargin = pageMargin?.Left?.Value ?? 0;
            uint rightMargin = pageMargin?.Right?.Value ?? 0;
            int topMargin = pageMargin?.Top?.Value ?? 0;
            int bottomMargin = pageMargin?.Bottom?.Value ?? 0;
            uint width = pageSize.Width.Value - leftMargin - rightMargin;
            int height = ((int)pageSize.Height.Value) - topMargin - bottomMargin;
            if (pageSize.Orient != null && pageSize.Orient.Value == PageOrientationValues.Landscape)
            {
                return new Size(height, (int)width);
            }
            else
            {
                return new Size((int)width, height);
            }
        }
        // If not found, return A4 portrait.
        // A4 page size is 210 mm x 297 mm = 11906 x 16838
        return new Size(11906, 16838);
    }

    public static string GetAvailablePartId(MainDocumentPart mainPart)
    {
        var rIds = mainPart.Parts.Where(part => part.RelationshipId.StartsWith("rId")).ToList();
        if (!rIds.Any())
            return "rId0";

        var maxId = rIds.Max(part => part.RelationshipId)?.TrimStart("rId");
        if (!string.IsNullOrWhiteSpace(maxId) && long.TryParse(maxId, NumberStyles.Integer, CultureInfo.InvariantCulture, out long id))
            return $"rId{id + 1}";
        else
            // Unexpected rId format, return a random ID
            return $"rId{new Random().Next(1000, 999999999)}";
    }
}
