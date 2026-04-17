using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class DrawingHelpers
{
    public static bool IsInline(this Drawing drawing)
    {
        return drawing.Inline != null;
    }

    public static bool IsFloating(this Drawing drawing)
    {
        return drawing.Inline == null &&
                  drawing.Anchor?.GetFirstChild<WrapTopBottom>() == null &&
                  drawing.Anchor?.GetFirstChild<WrapSquare>() == null &&
                  drawing.Anchor?.GetFirstChild<WrapTight>() == null &&
                  drawing.Anchor?.GetFirstChild<WrapThrough>() == null;
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

    // public static bool IsInline(this Picture picture)
    // {
    // }

    // public static bool IsFloating(this Picture picture)
    // {
    // }

    // internal static bool IsLayoutSupported(this Picture picture, ImageLayoutType layoutType)
    // {
    // }
}
