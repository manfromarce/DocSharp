using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using Shadow14 = DocumentFormat.OpenXml.Office2010.Word.Shadow;
using Outline14 = DocumentFormat.OpenXml.Office2010.Word.TextOutlineEffect;
using DocSharp.Helpers;
using M = DocumentFormat.OpenXml.Math;
using DocSharp.Writers;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{
    internal override void ProcessText(Text text, RtfStringWriter sb)
    {
        sb.WriteRtfEscaped(text.InnerText);
    }

    internal void ProcessText(M.Text text, RtfStringWriter sb)
    {
        sb.WriteRtfEscaped(text.InnerText);
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, RtfStringWriter sb)
    {
        sb.Write("\\chpgn ");
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, RtfStringWriter sb)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value) &&
            !string.IsNullOrEmpty(symbolChar?.Font?.Value))
        {
            fonts.TryAddAndGetIndex(symbolChar.Font.Value, out int fontIndex);
            sb.Write('{');
            sb.Write($"\\f{fontIndex} ");
            sb.WriteRtfUnicodeChar(symbolChar.Char.Value);
            sb.Write('}');
        }
    }   
}
