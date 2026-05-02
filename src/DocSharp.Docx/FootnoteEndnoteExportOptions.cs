namespace DocSharp.Docx;

public enum FootnoteEndnoteExportOptions
{
    /// <summary>
    /// Ignore foonotes and endnotes
    /// </summary>
    None,
    /// <summary>
    /// Export footnotes at the end of the section (because HTML is not paginated), and endnotes at the end of section or document depending on DOCX settings (EndnoteDocumentWideProperties).
    /// </summary>
    DocumentSettings,
    /// <summary>
    /// Export footnotes and endnotes at the end of the document, regardless of the document settings.
    /// </summary>
    EndOfDocument,
    /// <summary>
    /// Export footnotes at the end of the section, and endnotes at the end of the document, regardless of the document settings.
    /// </summary>
    FootnotesEndOfSection_EndnotesEndOfDocument
}
