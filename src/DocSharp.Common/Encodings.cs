using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp;

public static class Encodings
{
    private static readonly Encoding _utf8NoBOM = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    public static Encoding UTF8NoBOM => _utf8NoBOM;
}
