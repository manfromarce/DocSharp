using System;
using System.Collections.Generic;
using System.Text;

namespace DocSharp.Rtf;

public class GroupEnd : IToken
{
    public TokenType Type => TokenType.GroupEnd;

    public override string ToString()
    {
        return "}";
    }
}
