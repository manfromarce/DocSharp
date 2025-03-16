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
            return string.Empty; // Don't add trailing slash, it would be intended as root path.
        }
        else
        {
            return Path.TrimEndingDirectorySeparator(url.Replace('\\', '/')) + "/";
        }
    }
}
