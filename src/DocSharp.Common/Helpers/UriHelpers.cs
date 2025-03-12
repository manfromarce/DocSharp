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
        return Path.TrimEndingDirectorySeparator(uri.Trim('"').Replace('\\', '/')) + "/";
    }
}
