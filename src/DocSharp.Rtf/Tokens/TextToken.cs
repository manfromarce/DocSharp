using System;
using System.Collections.Generic;
using System.Text;

namespace DocSharp.Rtf;

public class TextToken : IToken
{
    public string Value { get; set; } = string.Empty;
    public TokenType Type => TokenType.Text;

    public override string ToString()
    {
        return Value;
    }
}
