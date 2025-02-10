using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public abstract class DocxConverterBase
{
    static DocxConverterBase()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <summary>
    /// Convert a <see cref="WordprocessingDocument"/> to a string in the output format.
    /// </summary>
    /// <param name="inputDocument">The DOCX document to use.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(WordprocessingDocument inputDocument)
    {
        var sb = new StringBuilder();
        var body = inputDocument.MainDocumentPart?.Document.Body;
        if (body != null)
        {
            ProcessBody(body, sb);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Convert a DOCX <see cref="Stream"/> to a string in the output format.
    /// </summary>
    /// <param name="inputStream">The DOCX Stream to use.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(Stream inputStream)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            return ConvertToString(wordDocument);
        }
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(string inputFilePath)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputFilePath, false))
        {
            return ConvertToString(wordDocument);
        }
    }

    /// <summary>
    /// Convert a DOCX file to a string in the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <returns>A string in the output format</returns>
    public string ConvertToString(byte[] inputBytes)
    {
        using (var memoryStream = new MemoryStream(inputBytes))
        {
            using (var wordDocument = WordprocessingDocument.Open(memoryStream, false))
            {
                return ConvertToString(wordDocument);
            }
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(WordprocessingDocument inputDocument, string outputFilePath)
    {
        File.WriteAllText(outputFilePath, ConvertToString(inputDocument));
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputDocument">The WordprocessingDocument to use.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(WordprocessingDocument inputDocument, Stream outputStream)
    {
        using (var streamWriter = new StreamWriter(outputStream, leaveOpen: true))
        {
            streamWriter.Write(ConvertToString(inputDocument));
        }
    }
    
    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(string inputFilePath, string outputFilePath)
    {
        File.WriteAllText(outputFilePath, ConvertToString(inputFilePath));
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputFilePath">The DOCX file path.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(string inputFilePath, Stream outputStream)
    {
        using (var streamWriter = new StreamWriter(outputStream, leaveOpen: true))
        {
            streamWriter.Write(ConvertToString(inputFilePath));
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(Stream inputStream, string outputFilePath)
    {
        File.WriteAllText(outputFilePath, ConvertToString(inputStream));
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputStream">The input DOCX stream to use.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(Stream inputStream, Stream outputStream)
    {
        using (var streamWriter = new StreamWriter(outputStream, leaveOpen: true))
        {
            streamWriter.Write(ConvertToString(inputStream));
        }
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <param name="outputFilePath">The output file path.</param>
    public void Convert(byte[] inputBytes, string outputFilePath)
    {
        File.WriteAllText(outputFilePath, ConvertToString(inputBytes));
    }

    /// <summary>
    /// Convert a DOCX file to the output format.
    /// </summary>
    /// <param name="inputBytes">The DOCX file bytes.</param>
    /// <param name="outputStream">The output stream.</param>
    public void Convert(byte[] inputBytes, Stream outputStream)
    {
        using (var streamWriter = new StreamWriter(outputStream, leaveOpen: true))
        {
            streamWriter.Write(ConvertToString(inputBytes));
        }
    }

    internal virtual void ProcessBody(Body body, StringBuilder sb)
    {
        foreach (var element in body.Elements())
        {
            ProcessBodyElement(element, sb);
        }
    }

    internal virtual void ProcessBodyElement(OpenXmlElement element, StringBuilder sb)
    {
        ProcessCompositeElement(element, sb);
    }

    internal virtual void ProcessCompositeElement(OpenXmlElement element, StringBuilder sb)
    {
        switch (element)
        {
            case Paragraph paragraph:
                ProcessParagraph(paragraph, sb);
                break;
            case Table table:
                ProcessTable(table, sb);
                break;
            case BookmarkStart bookmark:
                ProcessBookmarkStart(bookmark, sb);
                break;
            case BookmarkEnd bookmarkEnd:
                ProcessBookmarkEnd(bookmarkEnd, sb);
                break;
            case SdtBlock sdtBlock:
                ProcessSdtBlock(sdtBlock, sb);
                break;
        }
    }

    internal virtual void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        foreach(var element in paragraph.Elements())
        {
            ProcessParagraphElement(element, sb);
        }
    }

    // Used for paragraphs and other composite elements
    internal virtual void ProcessParagraphElement(OpenXmlElement element, StringBuilder sb)
    {
        switch (element)
        {
            case Run run:
                ProcessRun(run, sb);
                break;
            case BookmarkStart bookmarkStart:
                ProcessBookmarkStart(bookmarkStart, sb);
                break;
            case BookmarkEnd bookmarkEnd:
                ProcessBookmarkEnd(bookmarkEnd, sb);
                break;
            case Hyperlink hyperlink:
                ProcessHyperlink(hyperlink, sb);
                break;
            case Picture picture:
                ProcessPicture(picture, sb);
                break;
            case Drawing drawing:
                ProcessDrawing(drawing, sb);
                break;
            case SdtRun sdtRun:
                ProcessSdtRun(sdtRun, sb);
                break;
            case SimpleField simpleField:
                ProcessSimpleField(simpleField, sb);
                break;
        }
    }

    internal virtual bool ProcessRunElement(OpenXmlElement? element, StringBuilder sb)
    {
        switch (element)
        {
            case Text textElement:
                ProcessText(textElement, sb);
                return true;
            case Picture picture:
                ProcessPicture(picture, sb);
                return true;
            case Drawing drawing:
                ProcessDrawing(drawing, sb);
                return true;
            case Break br:
                ProcessBreak(br, sb);
                return true;
            case CarriageReturn:
                // The behavior of a carriage return is the same to a break character with null type and clear attributes
                // (source: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.carriagereturn)
                ProcessBreak(new Break() { }, sb);
                return true;
            case TabChar:
                ProcessText(new Text("\t"), sb);
                return true;
            case NoBreakHyphen:
                ProcessText(new Text("\u2011"), sb);
                return true;
            case FieldChar fieldChar:
                ProcessFieldChar(fieldChar, sb);
                return true;
            case FieldCode fieldCode:
                ProcessFieldCode(fieldCode, sb);
                return true;
            case DayShort:
                ProcessText(new Text(DateTime.Now.ToString("dd")), sb);
                break;
            case DayLong:
                ProcessText(new Text(DateTime.Now.ToString("dddd")), sb);
                break;
            case MonthShort:
                ProcessText(new Text(DateTime.Now.ToString("MM")), sb);
                break;
            case MonthLong:
                ProcessText(new Text(DateTime.Now.ToString("MMMM")), sb);
                break;
            case YearShort:
                ProcessText(new Text(DateTime.Now.ToString("YY")), sb);
                break;
            case YearLong:
                ProcessText(new Text(DateTime.Now.ToString("YYYY")), sb);
                return true;
            case AlternateContent alternateContent:
                if (!ProcessRunElement(alternateContent.GetFirstChild<AlternateContentChoice>()?.FirstChild, sb))
                {
                    return ProcessRunElement(alternateContent.GetFirstChild<AlternateContentFallback>()?.FirstChild, sb);
                }
                break;
        }
        return false;
    }

    internal virtual void ProcessSimpleField(SimpleField field, StringBuilder sb)
    {
        foreach (var element in field.Elements())
        {
            ProcessParagraphElement(element, sb);
        }
    }

    internal virtual void ProcessSdtRun(SdtRun sdtRun, StringBuilder sb)
    {
        if (sdtRun.SdtContentRun != null)
        {
            foreach (var element in sdtRun.Elements())
            {
                switch (element)
                {
                    case BookmarkStart bookmarkStart:
                        ProcessBookmarkStart(bookmarkStart, sb);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        ProcessBookmarkEnd(bookmarkEnd, sb);
                        break;
                }
            }
            foreach (var element in sdtRun.SdtContentRun.Elements())
            {
                ProcessParagraphElement(element, sb);
            }
        }
    }

    internal virtual void ProcessSdtBlock(SdtBlock sdtBlock, StringBuilder sb)
    {
        if (sdtBlock.SdtContentBlock != null)
        {
            foreach (var element in sdtBlock.Elements())
            {
                switch (element)
                {
                    case BookmarkStart bookmarkStart:
                        ProcessBookmarkStart(bookmarkStart, sb);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        ProcessBookmarkEnd(bookmarkEnd, sb);
                        break;
                }
            }
            foreach (var element in sdtBlock.SdtContentBlock.Elements())
            {
                ProcessCompositeElement(element, sb);
            }
        }
    }

    internal abstract void ProcessRun(Run run, StringBuilder sb);
    internal abstract void ProcessBreak(Break @break, StringBuilder sb);
    internal abstract void ProcessBookmarkStart(BookmarkStart bookmarkStart, StringBuilder sb);
    internal abstract void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, StringBuilder sb);
    internal abstract void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb);
    internal abstract void ProcessDrawing(Drawing picture, StringBuilder sb);
    internal abstract void ProcessPicture(Picture picture, StringBuilder sb);
    internal abstract void ProcessTable(Table table, StringBuilder sb);
    internal abstract void ProcessText(Text text, StringBuilder sb);
    internal abstract void ProcessFieldChar(FieldChar field, StringBuilder sb);
    internal abstract void ProcessFieldCode(FieldCode field, StringBuilder sb);

}
