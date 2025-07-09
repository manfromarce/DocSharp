using System.Text;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public sealed class LaTexStringWriter : BaseStringWriter
{
    public LaTexStringWriter()
    {
        NewLine = "\n"; // Use LF by default for LaTex
    }
}
