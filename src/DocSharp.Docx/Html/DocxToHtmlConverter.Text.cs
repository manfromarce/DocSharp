using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToTextConverterBase
{
    internal override void ProcessText(Text text, StringBuilder sb)
    {
        string font = string.Empty;
        if (text.Parent is Run run)
        {
            var fonts = OpenXmlHelpers.GetEffectiveProperty<RunFonts>(run);
            font = fonts?.Ascii?.Value?.ToLowerInvariant() ?? string.Empty;
        }
        string t = text.InnerText;
        var stringInfo = new StringInfo(t);
        for (int i = 0; i < stringInfo.LengthInTextElements; i++)
        {
            string textElement = stringInfo.SubstringByTextElements(i, 1);
            HtmlHelpers.Append(textElement, font, sb);
        }
    }

    internal override void ProcessSymbolChar(SymbolChar symbolChar, StringBuilder sb)
    {
        if (!string.IsNullOrEmpty(symbolChar?.Char?.Value))
        {
            string hexValue = symbolChar?.Char?.Value!;
            if (hexValue.StartsWith("0x", StringComparison.OrdinalIgnoreCase) ||
                hexValue.StartsWith("&h", StringComparison.OrdinalIgnoreCase))
            {
                hexValue = hexValue.Substring(2);
            }
            string htmlEntity = string.Empty;
            if (int.TryParse(hexValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int decimalValue))
            {
                if (!string.IsNullOrEmpty(symbolChar?.Font?.Value))
                {
                    htmlEntity = FontConverter.ToUnicode(symbolChar!.Font!.Value!, (char)decimalValue);
                }
            }
            if (string.IsNullOrEmpty(htmlEntity)) // If htmlEntity is empty, use the original char code
            {
                htmlEntity = $"&#{decimalValue.ToStringInvariant()};";
            }
            sb.Append(htmlEntity);
        }
    }

    internal override void ProcessPageNumber(PageNumber pageNumber, StringBuilder sb)
    {
    }
}
