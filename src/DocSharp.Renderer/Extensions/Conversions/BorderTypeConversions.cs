using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Text;
using DocSharp.Renderer.Core;
using PeachPDF.PdfSharpCore.Drawing;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer
{
    internal static class BorderTypeConversions
    {
        public static XPen ToPen(this Word.BorderType border, XPen defaultIfNull = null)
        {
            if (border == null)
            {
                return defaultIfNull;
            }

            var color = border.Color.ToXColor();
            var width = border.Size.EpToPoint();
            var val = border.Val?.Value ?? Word.BorderValues.Single;
            var pen = new XPen(color, width);
            pen.UpdateStyle(val);
            return pen;
        }

        private static void UpdateStyle(this XPen pen, Word.BorderValues borderValue)
        {
            switch (true)
            {
                case true when borderValue == Word.BorderValues.Nil:
                case true when borderValue == Word.BorderValues.None:
                    pen.Color = XColors.Transparent;
                    pen.Width = 0;
                    break;
                case true when borderValue == Word.BorderValues.Single:
                case true when borderValue == Word.BorderValues.Thick:
                    pen.DashStyle = XDashStyle.Solid;
                    break;
                case true when borderValue == Word.BorderValues.Dotted:
                    pen.DashStyle = XDashStyle.Dot;
                    break;
                case true when borderValue == Word.BorderValues.DashSmallGap:
                case true when borderValue == Word.BorderValues.Dashed:
                    pen.DashStyle = XDashStyle.Dash;
                    break;
                case true when borderValue == Word.BorderValues.DotDash:
                    pen.DashStyle = XDashStyle.DashDot;
                    break;
                case true when borderValue == Word.BorderValues.DotDotDash:
                    pen.DashStyle = XDashStyle.DashDotDot;
                    break;
            }
        }
    }
}
