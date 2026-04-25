namespace DocSharp.Docx;

public enum HeaderFooterExportOptions
{
    /// <summary>
    /// Ignore headers and footers
    /// </summary>
    None,
    /// <summary>
    /// Export the header of the first section (first page header if specified, otherwise default header) 
    /// and the footer of the last section (in case of different footers for even and odd pages, tries to detect the last page number and export the corresponding footer, otherwise exports the default footer).
    /// </summary>
    FirstHeaderLastFooter,
    /// <summary>
    /// Export the first header (first page header if specified, otherwise default header) and the last footer in each section.
    /// </summary>
    FirstHeaderLastFooterPerSection
}
