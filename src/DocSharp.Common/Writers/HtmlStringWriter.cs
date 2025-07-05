using System.Text;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public class HtmlStringWriter : BaseStringWriter
{
    public HtmlStringWriter() 
    {
        NewLine = "\n"; // Use LF by default for HTML
    }
}
