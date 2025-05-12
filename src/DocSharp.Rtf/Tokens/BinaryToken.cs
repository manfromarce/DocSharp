using System;
using System.Collections.Generic;
using System.Text;

namespace DocSharp.Rtf;

public class BinaryToken : IToken
{
    public byte[] Value { get; set; } = [];
    public TokenType Type => TokenType.Text;
}
