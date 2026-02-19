using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Helpers;

public class UriHelpers
{
    public static string NormalizeBaseUri(string uri)
    {
        string url = uri.Trim('"');
        if (string.IsNullOrEmpty(url))
        {
            // Same directory as the Markdown file
            return "./"; 
        }
        else
        {
            return url.Replace('\\', '/').TrimEnd('/') + "/";
        }
    }
}
