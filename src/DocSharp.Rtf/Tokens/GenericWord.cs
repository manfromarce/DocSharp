using System;
using System.Collections.Generic;
using System.Text;

namespace DocSharp.Rtf;

public class GenericWord : ControlWord<int>
{
    public override string Name { get; }

    public GenericWord(string name, int value) : base(value)
    {
        Name = name;
    }
}
