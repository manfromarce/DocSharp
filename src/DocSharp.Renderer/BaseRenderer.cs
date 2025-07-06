using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using PeachPDF.PdfSharpCore.Pdf;
using PeachPDF.PdfSharpCore.Utils;

namespace DocSharp.Renderer;

public abstract class BaseRenderer
{

    /// <summary>
    /// The default FontResolver supports Windows, macOS and Linux.  
    /// For other platforms you will need to set a custom FontResolver, otherwise changing this property is not necessary. 
    /// </summary>
    public static FontResolver FontResolver
    {
        get
        {
            _fontResolver ??= new FontResolver();
            return _fontResolver;
        }
        set
        {
            _fontResolver = value;
        }
    }

    private static FontResolver? _fontResolver;
}
