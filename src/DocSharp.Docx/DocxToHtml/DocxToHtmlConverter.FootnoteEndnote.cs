using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    // Note: FootnoteReference and EndnoteReference are found inside runs in the document body,
    // while FootnoteReferenceMark and EndnoteReferenceMark are in runs in the footnote/endnote part itself. 
    // The formatting is already processed for the parent run, 
    // this overrides are just to avoid writing square brackets in HTML (the base writer adds them for Markdown/plain text).
    internal override void ProcessFootnoteReference(FootnoteReference footnoteReference, HtmlTextWriter sb)
    {
        if (this.ExportFootnotesEndnotes)
        {
            ProcessText(new Text($"{footnoteReference.GetFootnoteIdString()}"), sb);
        }
    }

    internal override void ProcessEndnoteReference(EndnoteReference endnoteReference, HtmlTextWriter sb)
    {
        if (this.ExportFootnotesEndnotes)
        {
            ProcessText(new Text($"{endnoteReference.GetEndnoteIdString()}"), sb);
        }
    }

    internal override void ProcessFootnotes(FootnotesPart? footnotesPart, HtmlTextWriter sb)
    {
        if (this.ExportFootnotesEndnotes)
        {
            sb.WriteStartElement("div");
            base.ProcessFootnotes(footnotesPart, sb);
            sb.WriteEndElement("div");
        }
    }

    internal override void ProcessEndnotes(EndnotesPart? endnotesPart, HtmlTextWriter sb)
    {
        if (this.ExportFootnotesEndnotes)
        {
            sb.WriteStartElement("div");
            base.ProcessEndnotes(endnotesPart, sb);
            sb.WriteEndElement("div");
        }
    }

    internal override void ProcessFootnoteReferenceMark(FootnoteReferenceMark footnoteReferenceMark, HtmlTextWriter sb)
    {
        // We don't need to check ExportFootnotesEndnotes because it's already called inside the Foonotes part.
        ProcessText(new Text($"{footnoteReferenceMark.GetFootnoteIdString()}"), sb);
    }

    internal override void ProcessEndnoteReferenceMark(EndnoteReferenceMark endnoteReferenceMark, HtmlTextWriter sb)
    {
        // We don't need to check ExportFootnotesEndnotes because it's already called inside the Endnotes part.
        ProcessText(new Text($"{endnoteReferenceMark.GetEndnoteIdString()}"), sb);
    }
}
