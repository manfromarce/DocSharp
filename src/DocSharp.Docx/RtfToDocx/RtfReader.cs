using System.IO;
using System.Text;

namespace DocSharp.Rtf;

internal static class RtfReader
{
#if !NETFRAMEWORK
    static RtfReader()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }
#endif

    public static RtfDocument ReadRtf(TextReader reader)
    {
        var rtf = new RtfDocument();

        int i;
        char currentChar = default;
        char previousChar = default;
        while ((i = reader.Read()) != -1)
        {
            previousChar = currentChar;
            currentChar = (char)i;

            switch (currentChar)
            {
                case '{':
                    break;
                case '}':
                    break;
                case '\\':
                    break;
                case '*':
                    break;
                default:
                    if (char.IsDigit(currentChar))
                    {
                        
                    }
                    else if (IsEnglishLetter(currentChar))
                    {
                        
                    }
                    break;
            }
        }       

        return rtf;
    }

    private static bool IsEnglishLetter(char c)
    {
        return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
    }
}
