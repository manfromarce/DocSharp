using System;
using System.Collections.Generic;
using System.Text;
using DocSharp.Rtf.Tokens;

namespace DocSharp.Rtf;

public class ListStyleReference
{
    public int Id { get; }
    public ListStyle Style { get; }

    internal ListStyleReference(int id, ListStyle style)
    {
        Id = id;
        Style = style;
    }
}
