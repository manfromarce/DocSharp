using System.Drawing;
using System.Globalization;
using OpenXml = DocumentFormat.OpenXml;
using Word = DocumentFormat.OpenXml.Wordprocessing;

using static DocumentFormat.OpenXml.Wordprocessing.HighlightColorValues;
using PeachPDF.PdfSharpCore.Drawing;

namespace DocSharp.Renderer
{
    internal static class ColorConversions
    {
        public static Color ToColor(this Word.Highlight highlight)
        {
            var colorName = highlight?.Val ?? None;
            var c = colorName.Value.ToColor();
            return c;
        }

        public static XColor ToXColor(this Word.Highlight highlight)
        {
            var colorName = highlight?.Val ?? None;
            var c = colorName.Value.ToXColor();
            return c;
        }

        public static Color ToColor(this Word.Color color)
        {
            if(color == null)
            {
                return Color.Black;
            }

            return color.Val.ToColor();
        }

        public static XColor ToXColor(this Word.Color color)
        {
            if (color == null)
            {
                return XColors.Black;
            }

            return color.Val.ToXColor();
        }

        public static Color ToColor(this OpenXml.StringValue color)
        {
            var hex = color?.Value;
            var result = hex.ToColor();
            return result;
        }

        public static XColor ToXColor(this OpenXml.StringValue color)
        {
            var hex = color?.Value;
            var result = hex.ToXColor();
            return result;
        }

        private static Color ToColor(this string hex)
        {
            if (string.IsNullOrWhiteSpace(hex) || hex == "auto")
            {
                return Color.FromArgb(0, 0, 0);
            }

            var (r, g, b) = hex.ToRgb();
            return Color.FromArgb(r, g, b);
        }

        private static XColor ToXColor(this string hex)
        {
            if (string.IsNullOrWhiteSpace(hex) || hex == "auto")
            {
                return XColor.FromArgb(0, 0, 0);
            }

            var (r, g, b) = hex.ToRgb();
            return XColor.FromArgb(r, g, b);
        }

        private static (int r, int g, int b) ToRgb(this string hex)
        {
            var r = int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber);
            var g = int.Parse(hex.Substring(2, 2), NumberStyles.HexNumber);
            var b = int.Parse(hex.Substring(4, 2), NumberStyles.HexNumber);
            return (r, g, b);
        }

        private static Color ToColor(this Word.HighlightColorValues name)
        {
            return true switch
            {
                _ when name == Black => Color.FromArgb(0, 0, 0),
                _ when name == Blue => Color.FromArgb(0, 0, 0xFF),
                _ when name == Cyan => Color.FromArgb(0, 0xFF, 0xFF),
                _ when name == Green => Color.FromArgb(0, 0xFF, 0),
                _ when name == Magenta => Color.FromArgb(0xFF, 0, 0xFF),
                _ when name == Red => Color.FromArgb(0xFF, 0, 0),
                _ when name == Yellow => Color.FromArgb(0xFF, 0xFF, 0),
                _ when name == White => Color.FromArgb(0xFF, 0xFF, 0xFF),
                _ when name == DarkBlue => Color.FromArgb(0, 0, 0x80),
                _ when name == DarkCyan => Color.FromArgb(0, 0x80, 0x80),
                _ when name == DarkGreen => Color.FromArgb(0, 0x80, 0),
                _ when name == DarkMagenta => Color.FromArgb(0x80, 0, 0x80),
                _ when name == DarkRed => Color.FromArgb(0x80, 0, 0),
                _ when name == DarkYellow => Color.FromArgb(0x80, 0x80, 0),
                _ when name == DarkGray => Color.FromArgb(0x80, 0x80, 0x80),
                _ when name == LightGray => Color.FromArgb(0xC0, 0xC0, 0xC0),
                _ when name == None => Color.Empty,
                _ => Color.Empty
            };
        }

        private static XColor ToXColor(this Word.HighlightColorValues name)
        {
            return true switch
            {
                _ when name == Black => XColor.FromArgb(0, 0, 0),
                _ when name == Blue => XColor.FromArgb(0, 0, 0xFF),
                _ when name == Cyan => XColor.FromArgb(0, 0xFF, 0xFF),
                _ when name == Green => XColor.FromArgb(0, 0xFF, 0),
                _ when name == Magenta => XColor.FromArgb(0xFF, 0, 0xFF),
                _ when name == Red => XColor.FromArgb(0xFF, 0, 0),
                _ when name == Yellow => XColor.FromArgb(0xFF, 0xFF, 0),
                _ when name == White => XColor.FromArgb(0xFF, 0xFF, 0xFF),
                _ when name == DarkBlue => XColor.FromArgb(0, 0, 0x80),
                _ when name == DarkCyan => XColor.FromArgb(0, 0x80, 0x80),
                _ when name == DarkGreen => XColor.FromArgb(0, 0x80, 0),
                _ when name == DarkMagenta => XColor.FromArgb(0x80, 0, 0x80),
                _ when name == DarkRed => XColor.FromArgb(0x80, 0, 0),
                _ when name == DarkYellow => XColor.FromArgb(0x80, 0x80, 0),
                _ when name == DarkGray => XColor.FromArgb(0x80, 0x80, 0x80),
                _ when name == LightGray => XColor.FromArgb(0xC0, 0xC0, 0xC0),
                _ when name == None => XColor.Empty,
                _ => XColor.Empty
            };
        }
    }
}
