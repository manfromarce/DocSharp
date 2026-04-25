namespace DocSharp.Docx;

public enum FootnoteEndnoteExportOptions
{
    /// <summary>
    /// Ignore foonotes and endnotes
    /// </summary>
    None,
    /// <summary>
    /// Export footnotes and endnotes at the end of the document, regardless of the document settings.
    /// </summary>
    EndOfDocument,
}
