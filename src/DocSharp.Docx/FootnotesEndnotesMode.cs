namespace DocSharp.Docx;

public enum FootnotesEndnotesMode
{
    /// <summary>
    /// Footnotes and endnotes are not exported.
    /// </summary>
    None,
    /// <summary>
    /// Footnotes are exported at the end of the section they belong to, 
    /// endnotes are exported at the end of the section or document depending on their properties.
    /// </summary>
    Default,
    /// <summary>
    /// Footnotes and endnotes are exported at the end of the section they belong to.
    /// </summary>
    PerSection,
    /// <summary>
    /// Footnotes are exported at the end of the section they belong to, 
    /// endnotes are exported at the end of the document.
    /// </summary>
    FootnotesPerSectionEndnotesAtEnd,
    /// <summary>
    /// Footnotes and endnotes are exported at the end of the document.
    /// </summary>
    DocumentEnd, 
}
