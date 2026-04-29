using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using System.Linq;

namespace DocSharp.Docx;

public static class DrawingHelpers
{
    public static bool IsInline(this Drawing drawing)
    {
        return drawing.Inline != null;
    }

    public static bool IsFloating(this Drawing drawing)
    {
        return drawing.Anchor != null &&
               !drawing.Anchor.Elements<Wp.WrapTopBottom>().Any() &&
               !drawing.Anchor.Elements<Wp.WrapSquare>().Any() &&
               !drawing.Anchor.Elements<Wp.WrapTight>().Any() &&
               !drawing.Anchor.Elements<Wp.WrapThrough>().Any();
        // WrapNone = in front of / behind text
    }

    internal static bool IsLayoutSupported(this Drawing drawing, ImageLayoutType layoutType)
    {
        switch (layoutType)
        {
            case ImageLayoutType.None:
                return false;
            case ImageLayoutType.Inline:
                return drawing.IsInline();
            case ImageLayoutType.InlineAndAnchored:
                return !drawing.IsFloating();
            case ImageLayoutType.All: 
                return true;
        }
        return false;
    }
}
