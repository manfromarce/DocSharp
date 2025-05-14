using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using V = DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Office2019.Drawing.SVG;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxConverterBase
{
    internal override void ProcessDrawing(Drawing drawing, StringBuilder sb)
    {
        // DrawingML object or picture

        // Layout properties
        if (drawing.Inline != null)
        {

        }
        else if (drawing.Anchor != null)
        {
            if (!FixedLayout)
            {
                return;
            }

            var hPos = drawing.Anchor.HorizontalPosition;
            var vPos = drawing.Anchor.VerticalPosition;
            
            if (hPos != null)
            {
                var hRelativeFrom = hPos.RelativeFrom;
                var hPercentage = hPos.Descendants<DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentagePositionHeightOffset>();
                var hAlignment = hPos.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment>();
                var hOffset = hPos.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset>();
            }
            if (vPos != null)
            {
                var vRelativeFrom = vPos.RelativeFrom;
                var vPercentage = vPos.Descendants<DocumentFormat.OpenXml.Office2010.Word.Drawing.PercentagePositionHeightOffset>();
                var vAlignment = vPos.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalAlignment>();
                var vOffset = vPos.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.PositionOffset>();
            }

        }

        // Get width and height from drawing
        var extent = drawing.Descendants<Wp.Extent>().FirstOrDefault();

        // Actual drawing
        var graphicData = drawing.Descendants<A.GraphicData>().FirstOrDefault();
        if (graphicData != null && extent?.Cx != null && extent?.Cy != null)
        {
            double width = extent.Cx.Value / 12700.0; // Convert EMUs to points
            double height = extent.Cy.Value / 12700.0;

            // GraphicData can contain many different types of elements.
            // For now, we support pictures and some shapes.

            if (graphicData.GetFirstChild<Pic.Picture>() is Pic.Picture pic)
            {
                if (pic.BlipFill != null && pic.BlipFill.Blip is A.Blip blip)
                {
                    ProcessPictureFill(blip, drawing, width, height, sb);
                }
            }
            else if (graphicData.GetFirstChild<Wps.WordprocessingShape>() is Wps.WordprocessingShape shape &&
                     shape.GetFirstChild<Wps.ShapeProperties>() is Wps.ShapeProperties shapePr &&
                     shapePr.GetFirstChild<A.PresetGeometry>() is A.PresetGeometry presetGeometry &&
                     presetGeometry.Preset != null)
            {
                // Generate SVG for shapes
                if (presetGeometry.Preset.Value == A.ShapeTypeValues.Ellipse)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<ellipse cx=\"{(width / 2).ToStringInvariant()}\" cy=\"{(height / 2).ToStringInvariant()}\" rx=\"{(width / 2).ToStringInvariant()}\" ry=\"{(height / 2).ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Rectangle ||
                         presetGeometry.Preset.Value == A.ShapeTypeValues.Round1Rectangle ||
                         presetGeometry.Preset.Value == A.ShapeTypeValues.Round2DiagonalRectangle ||
                         presetGeometry.Preset.Value == A.ShapeTypeValues.Round2SameRectangle ||
                         presetGeometry.Preset.Value == A.ShapeTypeValues.SnipRoundRectangle ||
                         presetGeometry.Preset.Value == A.ShapeTypeValues.Snip1Rectangle ||
                         presetGeometry.Preset.Value == A.ShapeTypeValues.Snip2DiagonalRectangle ||
                         presetGeometry.Preset.Value == A.ShapeTypeValues.Snip2SameRectangle)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<rect x=\"0\" y=\"0\" width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.RoundRectangle)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<rect x=\"0\" y=\"0\" width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" rx=\"10\" ry=\"10\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Triangle)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{(width / 2).ToStringInvariant()},0 {width.ToStringInvariant()},{height.ToStringInvariant()} 0,{height.ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.RightTriangle)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"0,0 0,{height.ToStringInvariant()} {width.ToStringInvariant()},{height.ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Parallelogram)
                {
                    double offset = width * 0.2;
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{offset.ToStringInvariant()},0 {width.ToStringInvariant()},0 {(width - offset).ToStringInvariant()},{height.ToStringInvariant()} 0,{height.ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Diamond)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{(width / 2).ToStringInvariant()},0 {width.ToStringInvariant()},{(height / 2).ToStringInvariant()} {(width / 2).ToStringInvariant()},{height.ToStringInvariant()} 0,{(height / 2).ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Trapezoid ||  
                         presetGeometry.Preset.Value == A.ShapeTypeValues.NonIsoscelesTrapezoid)
                {
                    double topBase = width * 0.6;
                    double bottomBase = width;
                    double offsetX = (width - topBase) / 2;
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{offsetX.ToStringInvariant()},0 {(offsetX + topBase).ToStringInvariant()},0 {width.ToStringInvariant()},{height.ToStringInvariant()} 0,{height.ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Pentagon)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetRegularPolygonPoints(5, width, height)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Hexagon)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{(width / 4).ToStringInvariant()},0 {(width * 3 / 4).ToStringInvariant()},0 {width.ToStringInvariant()},{(height / 2).ToStringInvariant()} {(width * 3 / 4).ToStringInvariant()},{height.ToStringInvariant()} {(width / 4).ToStringInvariant()},{height.ToStringInvariant()} 0,{(height / 2).ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Heptagon)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetRegularPolygonPoints(7, width, height)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Octagon)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetRegularPolygonPoints(8, width, height)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Decagon)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetRegularPolygonPoints(10, width, height)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Dodecagon)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetRegularPolygonPoints(12, width, height)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Star4)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetStarPolygonPoints(4, width, height, 0.5)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Star5)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetStarPolygonPoints(5, width, height, 0.5)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Star6)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetStarPolygonPoints(6, width, height, 0.5)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Star7)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetStarPolygonPoints(7, width, height, 0.5)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Star8)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetStarPolygonPoints(8, width, height, 0.5)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Star10)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetStarPolygonPoints(10, width, height, 0.5)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Star12)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetStarPolygonPoints(12, width, height, 0.5)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Star16)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetStarPolygonPoints(16, width, height, 0.5)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Star24)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetStarPolygonPoints(16, width, height, 0.5)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.Star32)
                {
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{GetStarPolygonPoints(16, width, height, 0.5)}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.RightArrow)
                {
                    double h2 = height / 2;
                    double h3 = height * 3 / 4;
                    double arrowBody = width * 0.6;
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"0,{(h2 / 2).ToStringInvariant()} {arrowBody.ToStringInvariant()},{(h2 / 2).ToStringInvariant()} {arrowBody.ToStringInvariant()},0 {width.ToStringInvariant()},{h2.ToStringInvariant()} {arrowBody.ToStringInvariant()},{height.ToStringInvariant()} {arrowBody.ToStringInvariant()},{h3.ToStringInvariant()} 0,{h3.ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.LeftArrow)
                {
                    double h2 = height / 2;
                    double h3 = height * 3 / 4;
                    double arrowBody = width * 0.6;
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{width.ToStringInvariant()},{(h2 / 2).ToStringInvariant()} {(width - arrowBody).ToStringInvariant()},{(h2 / 2).ToStringInvariant()} {(width - arrowBody).ToStringInvariant()},0 0,{h2.ToStringInvariant()} {(width - arrowBody).ToStringInvariant()},{height.ToStringInvariant()} {(width - arrowBody).ToStringInvariant()},{h3.ToStringInvariant()} {width.ToStringInvariant()},{h3.ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.UpArrow)
                {
                    double w2 = width / 2;
                    double w3 = width * 3 / 4;
                    double arrowBody = height * 0.6;
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{w2.ToStringInvariant()},0 {width.ToStringInvariant()},{(height - arrowBody).ToStringInvariant()} {w3.ToStringInvariant()},{(height - arrowBody).ToStringInvariant()} {w3.ToStringInvariant()},{height.ToStringInvariant()} {(w2 / 2).ToStringInvariant()},{height.ToStringInvariant()} {(w2 / 2).ToStringInvariant()},{(height - arrowBody).ToStringInvariant()} 0,{(height - arrowBody).ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.DownArrow)
                {
                    double w2 = width / 2;
                    double w3 = width * 3 / 4;
                    double arrowBody = height * 0.6;
                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{w2.ToStringInvariant()},{height.ToStringInvariant()} {width.ToStringInvariant()},{(arrowBody).ToStringInvariant()} {w3.ToStringInvariant()},{(arrowBody).ToStringInvariant()} {w3.ToStringInvariant()},0 {(w2 / 2).ToStringInvariant()},0 {(w2 / 2).ToStringInvariant()},{(arrowBody).ToStringInvariant()} 0,{(arrowBody).ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.LeftRightArrow)
                {
                    //x1 = y1 = 0;
                    double x2 = width / 4;
                    double x3 = width * 3 / 4;
                    double x4 = width;

                    double y2 = height / 4;
                    double y3 = height / 2;
                    double y4 = height * 3 / 4;
                    double y5 = height;

                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"0,{y3.ToStringInvariant()} {x2.ToStringInvariant()},0 {x2.ToStringInvariant()},{y2.ToStringInvariant()} {x3.ToStringInvariant()},{y2.ToStringInvariant()} {x3.ToStringInvariant()},0 {x4.ToStringInvariant()},{y3.ToStringInvariant()} {x3.ToStringInvariant()},{y5.ToStringInvariant()} {x3.ToStringInvariant()},{y4.ToStringInvariant()} {x2.ToStringInvariant()},{y4.ToStringInvariant()} {x2.ToStringInvariant()},{y5.ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
                else if (presetGeometry.Preset.Value == A.ShapeTypeValues.UpDownArrow)
                {
                    double y2 = height / 4;
                    double y3 = height * 3 / 4;
                    double y4 = height;

                    double x2 = width / 4;
                    double x3 = width / 2;
                    double x4 = width * 3 / 4;
                    double x5 = width;

                    sb.Append($"<svg width=\"{width.ToStringInvariant()}\" height=\"{height.ToStringInvariant()}\" viewBox=\"0 0 {width.ToStringInvariant()} {height.ToStringInvariant()}\" xmlns=\"http://www.w3.org/2000/svg\">");
                    sb.Append($"<polygon points=\"{x3.ToStringInvariant()},0 0,{y2.ToStringInvariant()} {x2.ToStringInvariant()},{y2.ToStringInvariant()} {x2.ToStringInvariant()},{y3.ToStringInvariant()} 0,{y3.ToStringInvariant()} {x3.ToStringInvariant()},{y4.ToStringInvariant()} {x5.ToStringInvariant()},{y3.ToStringInvariant()} {x4.ToStringInvariant()},{y3.ToStringInvariant()} {x4.ToStringInvariant()},{y2.ToStringInvariant()} {x5.ToStringInvariant()},{y2.ToStringInvariant()}\"");
                    ProcessFill(shapePr, drawing, width, height, sb);
                    sb.Append("/>");
                    sb.AppendLine("</svg>");
                }
            }
        }
    }

    private static string GetRegularPolygonPoints(int sides, double width, double height)
    {
        double cx = width / 2;
        double cy = height / 2;
        double rx = width / 2;
        double ry = height / 2;
        double angleOffset = -Math.PI / 2;

        var points = new List<string>();
        for (int i = 0; i < sides; i++)
        {
            double angle = 2 * Math.PI * i / sides + angleOffset;
            double x = cx + rx * Math.Cos(angle);
            double y = cy + ry * Math.Sin(angle);
            points.Add($"{x.ToStringInvariant()},{y.ToStringInvariant()}");
        }
        return string.Join(" ", points);
    }

    private static string GetStarPolygonPoints(int points, double width, double height, double innerRadiusRatio = 0.5)
    {
        double cx = width / 2;
        double cy = height / 2;
        double rx = width / 2;
        double ry = height / 2;
        double angleOffset = -Math.PI / 2; // Start from top

        var result = new List<string>();
        int totalPoints = points * 2;
        for (int i = 0; i < totalPoints; i++)
        {
            double angle = angleOffset + i * Math.PI / points;
            double radiusX = (i % 2 == 0) ? rx : rx * innerRadiusRatio;
            double radiusY = (i % 2 == 0) ? ry : ry * innerRadiusRatio;
            double x = cx + radiusX * Math.Cos(angle);
            double y = cy + radiusY * Math.Sin(angle);
            result.Add($"{x.ToStringInvariant()},{y.ToStringInvariant()}");
        }
        return string.Join(" ", result);
    }

    private void ProcessFill(Wps.ShapeProperties shapePr, Drawing drawing, double width, double height, StringBuilder sb)
    {
        if (shapePr.GetFirstChild<A.BlipFill>() is A.BlipFill blipFill &&
            blipFill.Blip is A.Blip blip)
        {
            // Picture fill
            ProcessPictureFill(blip, drawing, width, height, sb);
        }
        else if (shapePr.GetFirstChild<A.SolidFill>() is A.SolidFill solidFill)
        {
            // Solid fill
            var color = OpenXmlHelpers.GetColor(solidFill);
            if (color != null)
            {
                sb.Append($" fill=\"{color}\"/>");
            }
        }
        else if (shapePr.GetFirstChild<A.GradientFill>() is A.GradientFill gradientFill)
        {
            // Gradient fill
        }
        else if (shapePr.GetFirstChild<A.PatternFill>() is A.PatternFill patternFill)
        {
            // Pattern fill
        }
    }

    internal void ProcessPictureFill(A.Blip blip, Drawing drawing, double width, double height, StringBuilder sb)
    {
        var mainDocumentPart = OpenXmlHelpers.GetMainDocumentPart(drawing);
        if (blip.Descendants<SVGBlip>().FirstOrDefault() is SVGBlip svgBlip &&
            svgBlip.Embed?.Value is string svgRelId)
        {
            // Prefer the actual SVG image as web browsers can display it.
            ProcessImagePart(mainDocumentPart, svgRelId, width, height, sb);
        }
        else if (blip.Embed?.Value is string relId)
        {
            ProcessImagePart(mainDocumentPart, relId, width, height, sb);
        }
    }
}
