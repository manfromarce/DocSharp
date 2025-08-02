using System;
using System.Collections.Generic;
using System.Text;

namespace DocSharp.Rtf;

public interface IWord : IToken
{
    string Name { get; }
}
