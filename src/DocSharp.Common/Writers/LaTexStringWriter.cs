using System.Text;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public class LaTexStringWriter : BaseStringWriter
{
    public LaTexStringWriter()
    {
        NewLine = "\n"; // Use LF by default for LaTex
    }
}
