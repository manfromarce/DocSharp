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
    internal override void ProcessText(Text text, QuestPdfModel output)
    {
        if (currentSpan.Count > 0 && !string.IsNullOrEmpty(text.Text))
            currentSpan.Peek().Text += text.Text;
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, QuestPdfModel output)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value) &&
            !string.IsNullOrEmpty(symbolChar?.Font?.Value))
        {
            // Parse the hex char code to a decimal code
            string hexValue = symbolChar?.Char?.Value!;
            if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                hexValue = hexValue.Substring(2);
            }
            if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int decimalValue))
            {
                if (currentRunContainer.Count > 0 && 
                    currentSpan.Count > 0) // SymbolChar can only be present inside a Run, just like regular Text elements.
                {
                    // Close and retrieve the current span
                    var oldSpan = currentSpan.Pop();

                    // Create a new span for the symbol with the specified font and char.
                    // The SymbolChar in DOCX has the same properties (bold, italic, color, ...) as the parent run, 
                    // except for the font family.
                    var symbolSpan = oldSpan.CloneEmpty();
                    symbolSpan.FontFamily = symbolChar!.Font!.Value!;
                    symbolSpan.Text = ((char)decimalValue).ToString(); // convert decimal char code to string.

                    // Add the new span to the paragraph/hyperlink.
                    currentRunContainer.Peek().AddSpan(symbolSpan);

                    // The old span was closed ahead of time to process the SymbolChar element.
                    // Create a new span with the same properties to contain further text elements. 
                    // The new span will be closed by the ProcessRun method.
                    // If there are no remaining elements, the new span will be empty 
                    // and will be ignored during rendering.
                    var newSpan = oldSpan.CloneEmpty();
                    currentSpan.Push(newSpan);
                }
            }
        }
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, QuestPdfModel output)
    {
        if (currentRunContainer.Count > 0 && 
            currentSpan.Count > 0) // PageNumber can only be present inside a Run, just like regular Text elements.
        {
            // Close and retrieve the current span
            var oldSpan = currentSpan.Pop();

            // Add a new QuestPdfPageNumber object to the current run container.
            currentRunContainer.Peek().AddPageNumber();

            // The old span was closed ahead of time to process the PageNumber element.
            // Create a new span with the same properties to contain further text elements. 
            // The new span will be closed by the ProcessRun method.
            // If there are no remaining elements, the new span will be empty 
            // and will be ignored during rendering.
            var newSpan = oldSpan.CloneEmpty();
            currentSpan.Push(newSpan);
        }
    }

    internal override void ProcessFieldChar(FieldChar field, QuestPdfModel output)
    {        
    }

    internal override void ProcessFieldCode(FieldCode field, QuestPdfModel output)
    {        
    }
}