using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Docx;

public enum HeadersFootersMode
{
    /// <summary>
    /// Headers and footers are not exported.
    /// </summary>
    None,
    /// <summary>
    /// Primary header (or first page header, if present) of the first section 
    /// is exported at the beginning of the document, 
    /// and primary footer of the last section (including linked to previous) 
    /// is exported at the end of the document.
    /// </summary>
    FirstSectionHeaderLastSectionFooter,
    /// <summary>
    /// Primary header (or first page header, if present) and footer 
    /// are exported at the beginning and the end of each section.
    /// </summary>
    PerSection
}
