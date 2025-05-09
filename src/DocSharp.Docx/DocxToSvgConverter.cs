using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public class DocxToSvgConverter : DocxConverterBase
{
    /// <summary>
    /// Image converter to preserve TIFF, EMF and other image types unsupported by browsers. 
    /// If the DocSharp.ImageSharp or DocSharp.SystemDrawing package is installed, 
    /// this property can be set to a new instance of ImageSharpConverter or SystemDrawingConverter. 
    /// </summary>
    public IImageConverter? ImageConverter { get; set; } = null;

    internal override void ProcessDocument(Document document, StringBuilder sb)
    {
        sb.Append("<svg xmlns=\"http://www.w3.org/2000/svg\" version=\"1.1\" width=\"100%\" height=\"100%\">");
        // Process body content
        base.ProcessDocument(document, sb);
        sb.Append("</svg>");
    }

    internal override void ProcessBookmarkEnd(BookmarkEnd bookmarkEnd, StringBuilder sb)
    {
        
    }

    internal override void ProcessBookmarkStart(BookmarkStart bookmarkStart, StringBuilder sb)
    {
        
    }

    internal override void ProcessBreak(Break @break, StringBuilder sb)
    {
        
    }

    internal override void ProcessContinuationSeparatorMark(ContinuationSeparatorMark continuationSepMark, StringBuilder sb)
    {
        
    }

    internal override void ProcessDocumentBackground(DocumentBackground background, StringBuilder sb)
    {
        
    }

    internal override void ProcessDrawing(Drawing picture, StringBuilder sb)
    {
        
    }

    internal override void ProcessEmbeddedObject(EmbeddedObject obj, StringBuilder sb)
    {
        
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, StringBuilder sb)
    {
        
    }

    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, StringBuilder sb)
    {
        
    }

    internal override void ProcessFieldChar(FieldChar field, StringBuilder sb)
    {
        
    }

    internal override void ProcessFieldCode(FieldCode field, StringBuilder sb)
    {
        
    }

    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, StringBuilder sb)
    {
        
    }

    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark endnoteReferenceMark, StringBuilder sb)
    {
        
    }

    internal override void ProcessHyperlink(Hyperlink hyperlink, StringBuilder sb)
    {
        
    }

    internal override void ProcessMathElement(OpenXmlElement element, StringBuilder sb)
    {
        
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, StringBuilder sb)
    {
        
    }

    internal override void ProcessPicture(Picture picture, StringBuilder sb)
    {
        
    }

    internal override void ProcessPositionalTab(PositionalTab posTab, StringBuilder sb)
    {
        
    }

    internal override void ProcessRun(Run run, StringBuilder sb)
    {
        
    }

    internal override void ProcessSeparatorMark(SeparatorMark separatorMark, StringBuilder sb)
    {
        
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, StringBuilder sb)
    {
        
    }

    internal override void ProcessTable(Table table, StringBuilder sb)
    {
        
    }

    internal override void ProcessText(Text text, StringBuilder sb)
    {
        
    }
}
