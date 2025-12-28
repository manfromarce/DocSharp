using System.Collections.Generic;

namespace DocSharp.Rtf;

internal class RtfGroup : RtfToken
{
    public List<RtfToken> Tokens { get; } = new List<RtfToken>();
    public RtfGroup() { }
}

