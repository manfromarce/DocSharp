using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Binary.OpenXmlLib;

public enum WordprocessingDocumentType
{
    /// <summary>
    /// Word Document (*.docx)
    /// </summary>
    Document,
    /// <summary>
    /// Word Template (*.dotx)
    /// </summary>
    MacroEnabledDocument,
    /// <summary>
    /// Word Macro-Enabled Document (*.docm)
    /// </summary>
    MacroEnabledTemplate,
    /// <summary>
    /// Word Macro-Enabled Template (*.dotm)
    /// </summary>
    Template
}

public enum SpreadsheetDocumentType
{
    /// <summary>
    /// Excel Workbook (*.xlsx)
    /// </summary>
    Workbook,
    /// <summary>
    /// Excel Template (*.xltx)
    /// </summary>
    Template,
    /// <summary>
    /// Excel Macro-Enabled Workbook (*.xlsm)
    /// </summary>
    MacroEnabledWorkbook,
    /// <summary>
    /// Excel Macro-Enabled Template (*.xltm)
    /// </summary>
    MacroEnabledTemplate,
}

public enum PresentationDocumentType
{
    /// <summary>
    /// PowerPoint Presentation (*.pptx)
    /// </summary>
    Presentation,
    /// <summary>
    /// PowerPoint Template (*.potx)
    /// </summary>
    Template,
    /// <summary>
    /// PowerPoint Show (*.ppsx)
    /// </summary>
    Slideshow,
    /// <summary>
    /// PowerPoint Macro-Enabled Presentation (*.pptm)
    /// </summary>
    MacroEnabledPresentation,
    /// <summary>
    /// PowerPoint Macro-Enabled Template (*.potm)
    /// </summary>
    MacroEnabledTemplate,
    /// <summary>
    /// PowerPoint Macro-Enabled Show (*.ppsm)
    /// </summary>    
    MacroEnabledSlideshow,
}
