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
    public static Body EnsureBody(this WordprocessingDocument wordprocessingDocument)
    {
        var mainPart = wordprocessingDocument.MainDocumentPart ?? wordprocessingDocument.AddMainDocumentPart();
        return mainPart.EnsureBody();
    }

    public static Body EnsureBody(this MainDocumentPart mainDocumentPart)
    {
        mainDocumentPart.Document ??= new Document();
        return mainDocumentPart.Document.EnsureBody();
    }

    public static Body EnsureBody(this Document document)
    {
        document.Body ??= new Body();
        return document.Body;
    }

    public static T? FirstOrDefault<T>(this OpenXmlElement element, Func<T, bool> condition) where T : OpenXmlElement
    {
        return element.Elements<T>().FirstOrDefault(condition);
    }

    public static bool EndsWith<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        return element.HasChildren && element.LastChild is T;
    }

    public static void RemoveAll<T>(this OpenXmlElement element, Func<T, bool> condition) where T : OpenXmlElement
    {
        foreach (var subElement in element.Elements<T>())
        {
            if (condition.Invoke(subElement))
            {
                subElement.Remove();
            }
        }
    }

    public static void RemoveAll<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        foreach (var subElement in element.Elements<T>())
        {
            subElement.Remove();
        }
    }

    public static void RemoveEmpty<T>(this OpenXmlElement element) where T : OpenXmlElement
    {
        foreach (var subElement in element.Elements<T>())
        {
            if (!subElement.HasChildren)
                subElement.Remove();
        }
    }   

    /// <summary>
    /// Set child element of the specified type. 
    /// Useful when a strongly-typed property is not provided by the Open XML SDK.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="element"></param>
    public static T SetElement<T>(this OpenXmlElement element, T value) where T : OpenXmlElement
    {        
        element.RemoveAll<T>();
        return element.AppendChild(value);
    }

    public static void Clear(this OpenXmlElement element)
    {
        element.RemoveAllChildren();
        element.ClearAllAttributes();
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
    /// Helper function to retrieve Theme part from an Open XML element.
    /// </summary>
    /// <returns></returns>
    public static DocumentFormat.OpenXml.Drawing.Theme? GetThemePart(this OpenXmlElement element)
    {
        return GetMainDocumentPart(element)?.ThemePart?.Theme;
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
