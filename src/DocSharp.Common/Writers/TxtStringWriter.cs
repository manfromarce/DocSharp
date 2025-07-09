using System;
using System.Text;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public sealed class TxtStringWriter : BaseStringWriter
{
    public void WriteText(string text, string fontName, TxtStringWriter sb)
    {
        foreach (char c in text)
        {
            sb.Write(FontConverter.ToUnicode(fontName, c));
        }
    }
}
