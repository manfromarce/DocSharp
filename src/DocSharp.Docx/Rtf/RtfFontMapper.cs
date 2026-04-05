using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Rtf;

internal static class RtfFontMapper
{
    internal static string GetFontType(string fontFamily)
    {
        if (fontFamily.Equals("Times New Roman", StringComparison.OrdinalIgnoreCase) || 
            fontFamily.Equals("Cambria Math", StringComparison.OrdinalIgnoreCase) || 
            fontFamily.Equals("Book Antiqua", StringComparison.OrdinalIgnoreCase) || 
            fontFamily.Equals("Bookman Old Style", StringComparison.OrdinalIgnoreCase) || 
            fontFamily.Equals("Lucida Bright", StringComparison.OrdinalIgnoreCase) || 
            fontFamily.Equals("Lucida Fax", StringComparison.OrdinalIgnoreCase) || 
            fontFamily.Equals("Rockwell", StringComparison.OrdinalIgnoreCase) ||
            fontFamily.Equals("Rockwell Nova", StringComparison.OrdinalIgnoreCase) ||
            fontFamily.Equals("Georgia", StringComparison.OrdinalIgnoreCase) ||
            fontFamily.StartsWith("Georgia ", StringComparison.OrdinalIgnoreCase) ||
            fontFamily.Equals("Palatino", StringComparison.OrdinalIgnoreCase) ||
            fontFamily.StartsWith("Palatino ", StringComparison.OrdinalIgnoreCase))
        {
            return "froman";
        }
        else if (fontFamily.Equals("Arial", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.StartsWith("Arial ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Calibri", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.StartsWith("Calibri ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Aptos", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.StartsWith("Aptos ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Helvetica", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.StartsWith("Helvetica ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Verdana", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.StartsWith("Verdana ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Lato", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.StartsWith("Lato ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Lucida Sans", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Lucida Sans Unicode", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Microsoft Sans Serif", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.Equals("Inter", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Liberation Sans", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.StartsWith("Liberation Sans ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Open Sans", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.StartsWith("Open Sans ", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.Equals("Roboto", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.StartsWith("Roboto ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.StartsWith("Gill Sans ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Segoe UI", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.StartsWith("Segoe UI ", StringComparison.OrdinalIgnoreCase))
        {
            return "fswiss";
        }
        else if (fontFamily.Equals("Courier New", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.Equals("Consolas", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Lucida Console ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Lucida Sans Typewriter ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.StartsWith("Cascadia ", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Pica", StringComparison.OrdinalIgnoreCase))
        {
            return "fmodern";
        }
        else if (fontFamily.Equals("Cursive", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("French Script MT", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Lucida Calligraphy", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Lucida Handwriting", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Freestyle Script", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Ink Draft", StringComparison.OrdinalIgnoreCase) ||
                 fontFamily.Equals("Comic Sans MS", StringComparison.OrdinalIgnoreCase))
        {
            return "fscript";
        }
        else if (fontFamily.Equals("Old English", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.Equals("ITC Zapf Chancery", StringComparison.OrdinalIgnoreCase))
        {
            return "fdecor";
        }
        else if (fontFamily.Equals("Symbol", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.Equals("Wingdings", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.Equals("Wingdings 2", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.Equals("Wingdings 3", StringComparison.OrdinalIgnoreCase) || 
                 fontFamily.Equals("Webdings", StringComparison.OrdinalIgnoreCase))
        {
            return "ftech";
        }
        else if (fontFamily.Equals("Miriam", StringComparison.OrdinalIgnoreCase))
        {
            return "fbidi";
        }
        else
        {
            return "fnil";
        }
    }
}