using System;
using System.Collections.Generic;
using System.Text;

namespace DocSharp.Rtf;

public interface IToken
{
    TokenType Type { get; }
}
