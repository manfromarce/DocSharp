using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace DocSharp.Helpers;

public static class HtmlHelpers
{
    public static void Append(char c, string font, StringBuilder sb)
    {
        if (c == '\r')
        {
            // Ignore as it's usually followed by \n
        }
        else if (c == '\n')
        {
            sb.Append("<br>");
        }
        else
        {
            string s = FontConverter.ToUnicode(font, c);
            // Append escaped (the string may contain special chars such as <, >, &, ", ')
            sb.Append(WebUtility.HtmlEncode(s));
            //sb.Append(HttpUtility.HtmlEncode(s));
        }
    }

    public static void Append(string val, string font, StringBuilder sb)
    {
        if (val == "\r")
        {
            // Ignore as it's usually followed by \n
        }
        else if (val == "\n")
        {
            sb.Append("<br>");
        }
        else
        {
            string s = FontConverter.ToUnicode(font, val);
            // Append escaped (the string may contain special chars such as <, >, &, ", ')
            sb.Append(WebUtility.HtmlEncode(s));
        }
    }
}
