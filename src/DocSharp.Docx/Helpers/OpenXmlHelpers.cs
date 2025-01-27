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
        var stylesPart = GetStylesPart(paragraph);

        // Check paragraph properties
        T? propertyValue = paragraph.ParagraphProperties?.GetFirstChild<T>();
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
    public static T? GetEffectiveProperty<T>(this Run run) where T : OpenXmlElement
    {
        var stylesPart = GetStylesPart(run);

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
        var stylesPart = GetStylesPart(table);

        // Check table properties
        var tableProperties = table.GetFirstChild<TableProperties>();
        T? propertyValue = tableProperties?.GetFirstChild<T>();
        if (propertyValue != null)
        {
            return propertyValue;
        }

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
        var stylesPart = GetStylesPart(cell);

        // Check cell properties        
        T? propertyValue = cell.TableCellProperties?.GetFirstChild<T>();
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
