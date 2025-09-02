using System;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Markdig.Renderers.Docx;

internal class DocxTemplateHelper
{
    internal const string defaultTemplate = "DocSharp.Markdown.Docx.Resources.markdown-template.docx";

    internal static Stream LoadDefaultTemplate()
    {
        var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(defaultTemplate);
        if (stream == null)
        {
            stream = Assembly.GetCallingAssembly().GetManifestResourceStream(defaultTemplate);
        }
        if (stream == null)
        {
            throw new FileNotFoundException($"Failed to load default template from resources.");
        }
        return stream;
    }

    internal static WordprocessingDocument BuildFromDefaultTemplate(Stream outputStream, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document)
    {
        using (var stream = LoadDefaultTemplate())
        {
            stream.CopyTo(outputStream);
            var document = WordprocessingDocument.Open(outputStream, true);
            if (openXmlDocumentType != WordprocessingDocumentType.Document)
            {
                // This will create a template (.dotx) or macro-enabled document, if desired.
                document.ChangeDocumentType(openXmlDocumentType);
            }
            return document;
        }
    }

    internal static WordprocessingDocument BuildFromDefaultTemplate(string outputFilePath, WordprocessingDocumentType openXmlDocumentType = WordprocessingDocumentType.Document)
    {
        using (var stream = LoadDefaultTemplate())
        {
            using (var fs = new FileStream(outputFilePath, FileMode.Create, FileAccess.ReadWrite))
            {
                stream.CopyTo(fs);
            }

            var document = WordprocessingDocument.Open(outputFilePath, true);

            if (openXmlDocumentType != WordprocessingDocumentType.Document)
            {
                // This will create a template (.dotx) or macro-enabled document, if desired.
                document.ChangeDocumentType(openXmlDocumentType);
            }
            return document;
        }
    }

    internal static void CloneStyle(Style style, StyleDefinitionsPart sourceStylesPart, StyleDefinitionsPart targetStylesPart)
    {
        targetStylesPart.Styles ??= new Styles();
        if (style.StyleId?.Value is string styleId &&
            !targetStylesPart.Styles.Elements<Style>().Any(s => s.StyleId == styleId))
        {
            // Clone the basedOn style too, if not present in the target document
            var basedOn = style.BasedOn?.Val?.Value;
            if (basedOn != null &&
                sourceStylesPart.Styles != null && 
                !targetStylesPart.Styles.Elements<Style>().Any(s => s.StyleId == basedOn) &&
                sourceStylesPart.Styles.Elements<Style>().Any(s => s.StyleId == basedOn))
            {
                var basedOnStyle = sourceStylesPart.Styles.Elements<Style>()
                                                          .FirstOrDefault(s => s.StyleId == basedOn);
                if (basedOnStyle != null)
                {
                    CloneStyle(basedOnStyle, sourceStylesPart, targetStylesPart);
                }
            }

            // Clone the style
            targetStylesPart.Styles.Append(style.CloneNode(true));

            // Clone the "linked" and "next" styles too, if not present in the target document
            var linkedStyle = style.LinkedStyle?.Val?.Value;
            if (linkedStyle != null &&
                sourceStylesPart.Styles != null &&
                !targetStylesPart.Styles.Elements<Style>().Any(s => s.StyleId == linkedStyle) &&
                sourceStylesPart.Styles.Elements<Style>().Any(s => s.StyleId == linkedStyle))
            {
                var lStyle = sourceStylesPart.Styles.Elements<Style>()
                                                    .FirstOrDefault(s => s.StyleId == linkedStyle);
                if (lStyle != null)
                {
                    targetStylesPart.Styles.Append(lStyle.CloneNode(true));
                }
            }

            var nextStyle = style.NextParagraphStyle?.Val?.Value;
            if (nextStyle != null &&
                sourceStylesPart.Styles != null && 
                !targetStylesPart.Styles.Elements<Style>().Any(s => s.StyleId == nextStyle) &&
                sourceStylesPart.Styles.Elements<Style>().Any(s => s.StyleId == nextStyle))
            {
                var nStyle = sourceStylesPart.Styles.Elements<Style>()
                                                    .FirstOrDefault(s => s.StyleId == nextStyle);
                if (nStyle != null)
                {
                    targetStylesPart.Styles.Append(nStyle.CloneNode(true));
                }
            }
        }
    }

    internal static void AddStylesIfRequired(DocumentStyles styles, WordprocessingDocument targetDocument)
    {
        using (var templateStream = LoadDefaultTemplate())
        {
            using (WordprocessingDocument templateDocument = WordprocessingDocument.Open(templateStream, false))
            {
                if (targetDocument.MainDocumentPart is null)
                {
                    targetDocument.AddMainDocumentPart();
                }
                if (templateDocument.MainDocumentPart?.StyleDefinitionsPart is StyleDefinitionsPart templateStylesPart &&
                    templateStylesPart.Styles != null)
                {
                    var targetStylesPart = targetDocument.MainDocumentPart!.StyleDefinitionsPart;
                    targetStylesPart ??= targetDocument.MainDocumentPart?.AddNewPart<StyleDefinitionsPart>();
                    targetStylesPart!.Styles ??= new Styles();
                    foreach (Style style in templateStylesPart.Styles.Elements<Style>())
                    {
                        // Clone styles not defined in the target document
                        if (style.StyleId?.Value is string styleId &&
                            styles.Contains(styleId) &&
                            !targetStylesPart.Styles.Elements<Style>().Any(s => s.StyleId == styleId))
                        {
                            CloneStyle(style, templateStylesPart, targetStylesPart);

                            // Special handling for MDBulletedListItem and MDOrderedListItem: 
                            // - if we are adding these styles, we also need to add the associated numbering definitions.
                            // - if these styles are already present, use them as-is 
                            // (the bullet style or numbering type can be customized by editing the paragraph style;
                            // the user is responsible for ensuring that an associated numbering definition exists, 
                            // otherwise list items will be rendered as normal paragraphs).

                            // In the first case, to avoid conflicts we will create a new numbering instance
                            // and abstract numbering definition.

                            // Get the numbering instance associated to the style
                            if (styleId == "MDBulletedListItem" || styleId == "MDOrderedListItem")
                            {
                                var numPr = style.StyleParagraphProperties?.NumberingProperties?.NumberingId;
                                if (numPr?.Val != null && 
                                    templateDocument.MainDocumentPart?.NumberingDefinitionsPart is NumberingDefinitionsPart templateNumPart &&
                                    templateNumPart.Numbering?.Elements<NumberingInstance>()
                                                              .FirstOrDefault(ni => ni.NumberID != null &&  ni.NumberID == numPr.Val) 
                                                              is NumberingInstance templateNumInstance
                                                              && templateNumInstance.AbstractNumId?.Val != null)
                                {
                                    var templateAbstractNum = templateNumPart.Numbering.Elements<AbstractNum>()
                                                                 .FirstOrDefault(an => an.AbstractNumberId != null && 
                                                                                       an.AbstractNumberId == templateNumInstance.AbstractNumId.Val);
                                    if (templateAbstractNum != null)
                                    {
                                        // Ensure the target document has a numbering part
                                        var targetNumPart = targetDocument.MainDocumentPart!.NumberingDefinitionsPart ?? 
                                                            targetDocument.MainDocumentPart?.AddNewPart<NumberingDefinitionsPart>();
                                        targetNumPart!.Numbering ??= new Numbering();

                                        // Retrieve the next available numbering and abstract numbering IDs
                                        int nextAbstractNumId = 1;
                                        if (targetNumPart.Numbering.Elements<AbstractNum>().Any())
                                        {
                                            nextAbstractNumId = targetNumPart.Numbering.Elements<AbstractNum>()
                                                                         .Max(an => an.AbstractNumberId?.Value ?? 0) + 1;
                                        }
                                        int nextNumId = 1;
                                        if (targetNumPart.Numbering.Elements<NumberingInstance>().Any())
                                        {
                                            nextNumId = targetNumPart.Numbering.Elements<NumberingInstance>()
                                                                 .Max(ni => ni.NumberID?.Value ?? 0) + 1;
                                        }

                                        // Clone the abstract numbering definition with the new ID
                                        var newAbstractNum = (AbstractNum)templateAbstractNum.CloneNode(true);
                                        newAbstractNum.AbstractNumberId = nextAbstractNumId;
                                        // Insert after the last abstact numbering definition, before numbering instances
                                        // (not doing so causes issues with Word)
                                        var lastAbstractNum = targetNumPart.Numbering.Elements<AbstractNum>().LastOrDefault();
                                        if (lastAbstractNum != null)
                                            lastAbstractNum.InsertAfterSelf(newAbstractNum);
                                        else
                                            targetNumPart.Numbering.Append(newAbstractNum);

                                        // Create a new numbering instance pointing to the new abstract numbering definition
                                        var newNumInstance = new NumberingInstance { NumberID = nextNumId };
                                        newNumInstance.Append(new AbstractNumId() { Val = nextAbstractNumId });
                                        targetNumPart.Numbering.Append(newNumInstance);

                                        // Update the style to point to the new numbering definition
                                        var targetStyle = targetStylesPart.Styles.Elements<Style>()
                                                                            .FirstOrDefault(s => s.StyleId == styleId);
                                        if (targetStyle?.StyleParagraphProperties?.NumberingProperties?.NumberingId?.Val != null)
                                        {
                                            targetStyle.StyleParagraphProperties.NumberingProperties.NumberingId.Val = nextNumId;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
