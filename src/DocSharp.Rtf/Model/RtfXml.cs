using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocSharp.Rtf.Model;

internal class RtfXml
{
    public UnitValue DefaultTabWidth { get; set; }
    
    public Dictionary<IToken, object> Metadata { get; set; }
    
    public Element Root { get; set; }

    public Document RtfDocument { get; set; }
}
