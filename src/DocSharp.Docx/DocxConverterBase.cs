using System;
using System.IO;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public abstract class DocxConverterBase
{
    public string ConvertToString(Stream inputStream)
    {
        using (var wordDocument = WordprocessingDocument.Open(inputStream, false))
        {
            var sb = new StringBuilder();
            var body = wordDocument.MainDocumentPart?.Document.Body;
            if (body != null)
            {
                ProcessBody(body, sb);
            }
            return sb.ToString();
        }
    }

    public string ConvertToString(string inputFilePath)
    {
        using (var fileStream = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read))
        {
            return ConvertToString(fileStream);
        }
    }

    public void Convert(string inputFilePath, string outputFilePath)
    {
        File.WriteAllText(outputFilePath, ConvertToString(inputFilePath));
    }

    public void Convert(Stream inputStream, string outputStream)
    {
        using (var streamWriter = new StreamWriter(outputStream))
        {
            streamWriter.Write(ConvertToString(inputStream));
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
        switch (element)
        {
            case Paragraph paragraph:
                ProcessParagraph(paragraph, sb);
                break;           
            case Table table:
                ProcessTable(table, sb);
                break;
        }
    }

    internal virtual void ProcessParagraph(Paragraph paragraph, StringBuilder sb)
    {
        foreach(var element in paragraph.Elements())
        {
            switch (element)
            {
                case Run run:
                    ProcessRun(run, sb);
                    break;
                case Hyperlink hyperlink:
                    ProcessHyperlink(hyperlink, sb);
                    break;
                case Picture picture:
                    ProcessPicture(picture, sb);
                    break;
            }
        }
    }

    internal abstract void ProcessTable(Table table, StringBuilder sb);

    internal abstract void ProcessRun(Run run, StringBuilder sb);

    internal abstract void ProcessPicture(Picture picture, StringBuilder sb);

    internal abstract void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb);

}