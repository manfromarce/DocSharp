using System;
using System.Collections.Generic;
using System.Diagnostics;
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

namespace DocSharp.Docx;

public static class OpenXmlHelpers
{
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

    /// <summary>
    /// Helper function to retrieve main document part from an Open XML element.
    /// </summary>
    /// <returns></returns>
    public static MainDocumentPart? GetMainDocumentPart(this OpenXmlElement element)
    {
        var document = (element as Document) ?? element.Ancestors<Document>().FirstOrDefault();
        return document?.MainDocumentPart;
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

    // Helper function to get paragraph formatting from paragraph properties, style or default style.
    public static T? GetEffectiveProperty<T>(this Paragraph paragraph) where T : OpenXmlElement
    {
        // Check paragraph properties
        T? propertyValue = paragraph.ParagraphProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        var stylesPart = GetStylesPart(paragraph);

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
    public static T? GetEffectiveProperty<T>(this Run run) where T : OpenXmlElement
    {
        // Check run properties
        T? propertyValue = run.RunProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        // Check paragraph properties
        var paragraphProperties = run.GetFirstAncestor<Paragraph>()?.ParagraphProperties;
        propertyValue = paragraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        var stylesPart = GetStylesPart(run);

        // Check run style
        var runStyle = stylesPart.GetStyleFromId(run.RunProperties?.RunStyle?.Val, StyleValues.Character) ?? 
                       stylesPart.GetStyleFromId(run.RunProperties?.RunStyle?.Val, StyleValues.Paragraph);
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
        var paragraphStyle = stylesPart.GetStyleFromId(paragraphProperties?.ParagraphStyleId?.Val, StyleValues.Paragraph);
        while (paragraphStyle != null)
        {
            propertyValue = paragraphStyle.StyleRunProperties?.GetFirstChild<T>();
            if (propertyValue != null)
            {
                return propertyValue;
            }

            // Check styles from which the current style inherits
            paragraphStyle = stylesPart.GetBaseStyle(paragraphStyle);
        }

        // Check default run style for the current document
        return stylesPart.GetDefaultRunStyle()?.GetFirstChild<T>();
    }

    // Helper function to get table formatting from table properties or style.
    public static T? GetEffectiveProperty<T>(this Table table) where T : OpenXmlElement
    {
        // Check table properties
        var tableProperties = table.GetFirstChild<TableProperties>();
        T? propertyValue = tableProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

        var stylesPart = GetStylesPart(table);

        // Check table style
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        while (tableStyle != null)
        {
            propertyValue = tableStyle.StyleTableProperties?.GetFirstChild<T>();
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
    public static T? GetEffectiveProperty<T>(this TableRow row) where T : OpenXmlElement
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

        var stylesPart = GetStylesPart(row);

        // Check table style
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        while (tableStyle != null)
        {
            propertyValue = tableStyle.StyleTableProperties?.GetFirstChild<T>();
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
    public static T? GetEffectiveProperty<T>(this TableCell cell) where T : OpenXmlElement
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

        var stylesPart = GetStylesPart(cell);

        // Check table style
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        while (tableStyle != null)
        {
            propertyValue = tableStyle.StyleTableCellProperties?.GetFirstChild<T>() ?? 
                            tableStyle.StyleTableProperties?.GetFirstChild<T>();
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
                                                 int rowNumber, int columnNumber, int rowCount, int columnCount, bool isRightToLeft)
    {
        bool isFirstRow = rowNumber == 1;
        bool isFirstColumn = columnNumber == 1;
        bool isLastRow = rowNumber == rowCount;
        bool isLastColumn = columnNumber == columnCount;

        var targetTypesCell = new List<Type>();
        var targetTypesTable = new List<Type>();
        switch (borderValue)
        {
            case Primitives.BorderValue.Left:
                targetTypesCell.Add(typeof(LeftBorder));
                targetTypesCell.Add(isRightToLeft ? typeof(EndBorder) : typeof(StartBorder));
                if (isFirstColumn)
                {
                    targetTypesTable.Add(typeof(LeftBorder));
                    targetTypesTable.Add(isRightToLeft ? typeof(EndBorder) : typeof(StartBorder));
                }
                else
                {
                    targetTypesTable.Add(typeof(InsideVerticalBorder));
                }
                break;
            case Primitives.BorderValue.Right:
                targetTypesCell.Add(typeof(RightBorder));
                targetTypesCell.Add(isRightToLeft ? typeof(StartBorder) : typeof(EndBorder));
                if (isLastColumn)
                {
                    targetTypesTable.Add(typeof(RightBorder));
                    targetTypesTable.Add(isRightToLeft ? typeof(StartBorder) : typeof(EndBorder));
                }
                else
                {
                    targetTypesTable.Add(typeof(InsideVerticalBorder));
                }
                break;
            case Primitives.BorderValue.Top:
                targetTypesCell.Add(typeof(TopBorder));
                targetTypesTable.Add(isFirstRow ? typeof(TopBorder) : typeof(InsideHorizontalBorder));
                break;
            case Primitives.BorderValue.Bottom:
                targetTypesCell.Add(typeof(BottomBorder));
                targetTypesTable.Add(isLastRow ? typeof(BottomBorder) : typeof(InsideHorizontalBorder));
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
        var stylesPart = GetStylesPart(cell);
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        while (tableStyle != null)
        {
            res = tableStyle.StyleTableProperties?.TableBorders?.FirstOrDefault(element => targetTypesTable.Contains(element.GetType()));
            if (res != null)
            {
                return (BorderType)res;
            }

            // Check styles from which the current style inherits
            tableStyle = stylesPart.GetBaseStyle(tableStyle);
        }
        return null;
    }

    // Helper function to get a border (top, bottom, left, start, diagonal...) from cell/table/row properties or style.
    public static T? GetEffectiveBorder<T>(this TableRow row) where T : BorderType
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

        var stylesPart = GetStylesPart(row);

        // Check table style
        var tableStyle = stylesPart.GetStyleFromId(tableProperties?.TableStyle?.Val, StyleValues.Table);
        while (tableStyle != null)
        {
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

    public static void InsertAfterLastOfType(this OpenXmlCompositeElement parent, OpenXmlElement element)
    {
        var refElement = parent.Elements().LastOrDefault(e => e.GetType() == element.GetType());
        if (refElement == null)
        {
            parent.AppendChild(element);
        }
        else
        {
            parent.InsertAfter(element, refElement);
        }
    }           
}
