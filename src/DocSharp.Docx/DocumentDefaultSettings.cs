using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Docx;

public class DocumentDefaultSettings
{
    /// <summary>
    /// Font name. Default is Calibri.
    /// </summary>
    public string FontName { get; set; } = "Calibri";

    /// <summary>
    /// Font size in points. Default = 12.
    /// </summary>
    public int FontSize { get; set; } = 12;

    /// <summary>
    /// Space after paragraphs in points. Default = 8.
    /// </summary>
    public int SpaceAfterParagraph { get; set; } = 8;

    /// <summary>
    /// Space between lines in multiple of lines. Default = 1.15
    /// </summary>
    public decimal LineSpacing { get; set; } = 1.15m;
}
