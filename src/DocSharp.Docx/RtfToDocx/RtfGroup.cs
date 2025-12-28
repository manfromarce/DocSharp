using System.Collections.Generic;

namespace DocSharp.Rtf;

internal class RtfGroup : RtfControlWord
{
    internal IEnumerable<RtfToken> Tokens;

    public RtfGroup(string name) : base(name)
    {
        Tokens = [];
    }
}

