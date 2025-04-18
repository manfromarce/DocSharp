using System.Text;

namespace DocSharp.Rtf;

internal interface IRtfReader
{
    Encoding Encoding { get; set; }

    int Peek();
    int Read();
}
