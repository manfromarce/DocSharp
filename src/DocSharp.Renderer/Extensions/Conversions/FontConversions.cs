using System.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer
{
    internal static class FontConversions
    {
        public static XFontStyle BoldStyle(this Bold bold, XFontStyle defaultFontStyle)
        {
            return bold.OnOffTypeToStyle(XFontStyle.Bold, defaultFontStyle & XFontStyle.Bold);
        }

        public static XFontStyle ItalicStyle(this Italic italic, XFontStyle defaultFontStyle)
        {
            return italic.OnOffTypeToStyle(XFontStyle.Italic, defaultFontStyle & XFontStyle.Italic);
        }

        public static XFontStyle StrikeStyle(this Strike strike, XFontStyle defaultFontStyle)
        {
            return strike.OnOffTypeToStyle(XFontStyle.Strikeout, defaultFontStyle & XFontStyle.Strikeout);
        }

        public static XFontStyle UnderlineStyle(this Underline underline, XFontStyle defaultFontStyle)
        {
            if(underline?.Val == null)
            {
                return defaultFontStyle & XFontStyle.Underline;
            }

            return underline.Val.Value != UnderlineValues.None
                ? XFontStyle.Underline
                : XFontStyle.Regular;
        }

        public static double ToDouble(this FontSize? fontSize, double ifNull)
        {
            if(fontSize?.Val == null)
            {
                return ifNull;
            }

            var size = fontSize.Val.HPToPoint(ifNull);
            return size;
        }

        public static float ToFloat(this FontSize fontSize, float ifNull)
        {
            var size = fontSize.ToDouble(ifNull);
            return (float)size;
        }

        private static XFontStyle OnOffTypeToStyle(this OnOffType onOff, XFontStyle onValue, XFontStyle nullValue)
        {
            if(onOff == null)
            {
                return nullValue;
            }

            return (onOff.Val?.Value ?? true)
                ? onValue
                : XFontStyle.Regular;
        }
    }
}
