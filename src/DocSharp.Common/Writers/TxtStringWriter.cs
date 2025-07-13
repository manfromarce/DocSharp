using System;
using System.Text;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public sealed class TxtStringWriter : BaseStringWriter
{
    public void WriteText(string text, string fontName)
    {
        foreach (char c in text)
        {
            Write(FontConverter.ToUnicode(fontName, c));
        }
    }
}
