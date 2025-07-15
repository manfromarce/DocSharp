using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using O = DocumentFormat.OpenXml.Vml.Office;
using DocSharp.IO;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Globalization;

namespace DocSharp.Docx;

public partial class DocxToRtfConverter : DocxToTextConverterBase<RtfStringWriter>
{    
    internal override void ProcessDocumentBackground(DocumentBackground documentBackground, RtfStringWriter sb)
    {
        if (documentBackground.Background != null)
        {
            ProcessBackground(documentBackground.Background, sb);
        }
        else if (documentBackground.Color?.Value != null)
        {
            int? bgr = ColorHelpers.HexToBgr(documentBackground.Color);
            if (bgr != null)
            {
                sb.WriteLine(@"{\*\background {\shp{\*\shpinst\shpleft0\shptop0\shpright0\shpbottom0\shpfhdr0\shpbxmargin\shpbxignore\shpbymargin\shpbyignore\shpwr0\shpwrk0\shpfblwtxt1\shpz0\shplid1025");
                sb.WriteShapeProperty("shapeType", "1");
                sb.WriteShapeProperty("fFlipH", "0");
                sb.WriteShapeProperty("fFlipV", "0");
                sb.WriteShapeProperty("fillColor", bgr.Value);
                sb.WriteShapeProperty("fFilled", "1");
                sb.WriteShapeProperty("lineWidth", "0");
                sb.WriteShapeProperty("fLine", "0");
                sb.WriteShapeProperty("bWMode", "9");
                sb.WriteShapeProperty("fBackground", "1");
                sb.WriteShapeProperty("fLayoutInCell", "1");
                sb.WriteLine("}}}");

                sb.WriteLine(@"\viewbksp1");
            }
        }        
    }

    internal void ProcessBackground(V.Background background, RtfStringWriter sb)
    {
        if (background.Fill is V.Fill fill && fill.Type != null && fill.Type.HasValue)
        {
            sb.WriteLine(@"{\*\background {\shp{\*\shpinst\shpleft0\shptop0\shpright0\shpbottom0\shpfhdr0\shpbxmargin\shpbxignore\shpbymargin\shpbyignore\shpwr0\shpwrk0\shpfblwtxt1\shpz0\shplid1025");

            sb.WriteShapeProperty("shapeType", "1");
            sb.WriteShapeProperty("fFlipH", "0");
            sb.WriteShapeProperty("fFlipV", "0");
            sb.WriteShapeProperty("lineWidth", "0");
            sb.WriteShapeProperty("fLine", "0");
            sb.WriteShapeProperty("fBackground", "1");
            sb.WriteShapeProperty("fLayoutInCell", "1");

            if (background.Filled == null || (background.Filled != null && background.Filled.Value))
            {
                sb.WriteShapeProperty("fFilled", "1"); // Default
            }
            else
            {
                sb.WriteShapeProperty("fFilled", "0");
            }

            if (background.BlackWhiteMode != null && background.BlackWhiteMode.HasValue)
            {
                if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.Color)
                {
                    sb.WriteShapeProperty("bWMode", "0");
                }
                else if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.Auto)
                {
                    sb.WriteShapeProperty("bWMode", "1");
                }
                else if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.GrayScale)
                {
                    sb.WriteShapeProperty("bWMode", "2");
                }
                else if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.LightGrayScale)
                {
                    sb.WriteShapeProperty("bWMode", "3");
                }
                else if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.InverseGray)
                {
                    sb.WriteShapeProperty("bWMode", "4");
                }
                else if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.GrayOutline)
                {
                    sb.WriteShapeProperty("bWMode", "5");
                }
                else if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.BlackTextAndLines)
                {
                    sb.WriteShapeProperty("bWMode", "6");
                }
                else if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.HighContrast)
                {
                    sb.WriteShapeProperty("bWMode", "7");
                }
                else if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.Black)
                {
                    sb.WriteShapeProperty("bWMode", "8");
                }
                else if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.White)
                {
                    sb.WriteShapeProperty("bWMode", "9");
                }
                else if (background.BlackWhiteMode.Value == O.BlackAndWhiteModeValues.Undrawn)
                {
                    sb.WriteShapeProperty("bWMode", "10");
                }
            }

            if (background.PureBlackWhiteMode != null && background.PureBlackWhiteMode.HasValue)
            {
                if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.Color)
                {
                    sb.WriteShapeProperty("bWModePureBW", "0");
                }
                else if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.Auto)
                {
                    sb.WriteShapeProperty("bWModePureBW", "1");
                }
                else if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.GrayScale)
                {
                    sb.WriteShapeProperty("bWModePureBW", "2");
                }
                else if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.LightGrayScale)
                {
                    sb.WriteShapeProperty("bWModePureBW", "3");
                }
                else if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.InverseGray)
                {
                    sb.WriteShapeProperty("bWModePureBW", "4");
                }
                else if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.GrayOutline)
                {
                    sb.WriteShapeProperty("bWModePureBW", "5");
                }
                else if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.BlackTextAndLines)
                {
                    sb.WriteShapeProperty("bWModePureBW", "6");
                }
                else if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.HighContrast)
                {
                    sb.WriteShapeProperty("bWModePureBW", "7");
                }
                else if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.Black)
                {
                    sb.WriteShapeProperty("bWModePureBW", "8");
                }
                else if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.White)
                {
                    sb.WriteShapeProperty("bWModePureBW", "9");
                }
                else if (background.PureBlackWhiteMode.Value == O.BlackAndWhiteModeValues.Undrawn)
                {
                    sb.WriteShapeProperty("bWModePureBW", "10");
                }
            }

            if (background.NormalBlackWhiteMode != null && background.NormalBlackWhiteMode.HasValue)
            {
                if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.Color)
                {
                    sb.WriteShapeProperty("bWModeBW", "0");
                }
                else if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.Auto)
                {
                    sb.WriteShapeProperty("bWModeBW", "1");
                }
                else if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.GrayScale)
                {
                    sb.WriteShapeProperty("bWModeBW", "2");
                }
                else if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.LightGrayScale)
                {
                    sb.WriteShapeProperty("bWModeBW", "3");
                }
                else if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.InverseGray)
                {
                    sb.WriteShapeProperty("bWModeBW", "4");
                }
                else if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.GrayOutline)
                {
                    sb.WriteShapeProperty("bWModeBW", "5");
                }
                else if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.BlackTextAndLines)
                {
                    sb.WriteShapeProperty("bWModeBW", "6");
                }
                else if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.HighContrast)
                {
                    sb.WriteShapeProperty("bWModeBW", "7");
                }
                else if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.Black)
                {
                    sb.WriteShapeProperty("bWModeBW", "8");
                }
                else if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.White)
                {
                    sb.WriteShapeProperty("bWModeBW", "9");
                }
                else if (background.NormalBlackWhiteMode.Value == O.BlackAndWhiteModeValues.Undrawn)
                {
                    sb.WriteShapeProperty("bWModeBW", "10");
                }
            }

            int? bgr = ColorHelpers.HexToBgr(background.Fillcolor);
            if (bgr != null)
            {
                sb.WriteShapeProperty("fillColor", bgr.Value);
            }
            else
            {
                bgr = ColorHelpers.HexToBgr(background.Fill.Color);
                if (bgr != null)
                {
                    sb.WriteShapeProperty("fillColor", bgr.Value);
                }
            }

            int? bgr2 = ColorHelpers.HexToBgr(background.Fill.Color2, bgr);
            if (bgr2 != null)
            {
                sb.WriteShapeProperty("fillBackColor", bgr2.Value);
            }

            if (background.Fill.Colors?.Value != null)
            {
                var gradientColors = background.Fill.Colors.Value.Split(';');
                string shadeColors = "";
                foreach (var gradientColor in gradientColors)
                {
                    var properties = gradientColor.Split(' ');
                    if (properties.Length >= 2 && 
                        double.TryParse(properties[0].Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out double pos) 
                        && ColorHelpers.HexToBgr(properties[1].Trim()) is int color)
                    {
                        shadeColors += $"({color},{(long)Math.Round(pos * 65536)});";
                    }
                }
                shadeColors = shadeColors.TrimEnd(';');
                /*
                 * TEST: 
                0 #03d4a8;.25 #21d6e0;.75 #0087e6;1 #005cbf
                --> 8;4;(11064323,0);(14734881,16384);(15107840,49152);(12540928,65536)
                 What are the first two values?
                 */
                if (!string.IsNullOrEmpty(shadeColors))
                    sb.WriteShapeProperty("fillShadeColors", "");
            }

            var type = fill.Type.Value;
            var extendedProperties = fill.GetFirstChild<O.FillExtendedProperties>();
            // FillExtendedProperties has priority if present.
            if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Solid) || 
                type == V.FillTypeValues.Solid)
            {
                sb.WriteShapeProperty("fillType", "0");
            }
            else if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Pattern) ||
                type == V.FillTypeValues.Pattern)
            {
                sb.WriteShapeProperty("fillType", "1");
            }
            else if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Tile) ||
                type == V.FillTypeValues.Tile) // Texture
            {
                sb.WriteShapeProperty("fillType", "2");
            }
            else if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Frame) ||
                type == V.FillTypeValues.Frame) // Picture
            {
                sb.WriteShapeProperty("fillType", "3");
            }
            else if (extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.GradientUnscaled)
            {
                sb.WriteShapeProperty("fillType", "4");
            }
            else if (extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.GradientCenter)
            {
                sb.WriteShapeProperty("fillType", "5"); // Gradient from center to corners
            }
            else if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.GradientRadial) ||
               type == V.FillTypeValues.GradientRadial)
            {
                sb.WriteShapeProperty("fillType", "6"); // Radial gradient
            }
            else if ((extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Gradient) ||
               type == V.FillTypeValues.Gradient) 
            {
                sb.WriteShapeProperty("fillType", "7"); // Horizontal, vertical or diagonal gradient (uses fillAngle)
            }
            else if (extendedProperties?.Type != null && extendedProperties.Type.Value == O.FillValues.Background)
            {
                sb.WriteShapeProperty("fillType", "9"); // Use background fill
            }

            if (fill.Method != null && fill.Method.HasValue)
            {
                if (fill.Method.Value == V.FillMethodValues.Any)
                {
                    // Do't write fillShadeType
                }
                else if (fill.Method.Value == V.FillMethodValues.Linear)
                {
                    sb.WriteShapeProperty("fillShadeType", "1");
                }
                else if (fill.Method.Value == V.FillMethodValues.Linearsigma)
                {
                    sb.WriteShapeProperty("fillShadeType", "1073741835");
                }
                else if (fill.Method.Value == V.FillMethodValues.None)
                {
                    sb.WriteShapeProperty("fillShadeType", "0");
                }
                else if (fill.Method.Value == V.FillMethodValues.Sigma)
                {
                    sb.WriteShapeProperty("fillShadeType", "1073741826");
                }
            }

            if (fill.Angle != null && fill.Angle.HasValue)
            {
                var dec = fill.Angle.Value * 65536;
                sb.WriteShapeProperty("fillAngle", (long)Math.Round(dec));
            }

            if (fill.Focus?.Value != null)
            {
                string focus = fill.Focus.Value.TrimEnd('%');
                if (int.TryParse(focus, NumberStyles.Number, CultureInfo.InvariantCulture, out int v))
                    sb.WriteShapeProperty("fillFocus", focus);
            }

            if (fill.FocusPosition?.Value != null)
            {
                string focusPos = fill.FocusPosition.Value;
                string[] split = focusPos.Split(',');
                if (split.Length >= 2)
                {
                    string s1 = split[0].Trim();
                    string s2 = split[1].Trim();
                    if (!double.TryParse(s1, NumberStyles.Number, CultureInfo.InvariantCulture, out double leftRight))
                    {
                        if (s1 == string.Empty)
                        {
                            leftRight = 0; // Recognize formats such as ",1" which means the first value is 0
                        }
                    }
                    if (!double.TryParse(s2, NumberStyles.Number, CultureInfo.InvariantCulture, out double topBottom))
                    {
                        if (s2 == string.Empty)
                        {
                            topBottom = 0;
                        }
                    }
                    long val1 = (long)Math.Round(leftRight * 65536);
                    long val2 = (long)Math.Round(topBottom * 65536);
                    long width = 0;
                    long height = 0;

                    if (fill.FocusSize?.Value != null && !string.IsNullOrEmpty(fill.FocusSize.Value))
                    {
                        string focusSize = fill.FocusSize.Value;
                        string[] size = focusSize.Split(',');
                        if (size.Length >= 2 && 
                            double.TryParse(size[0].Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out double w) &&
                            double.TryParse(size[1].Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out double h))
                        {
                            width = (long)Math.Round(w * 65536);
                            height = (long)Math.Round(h * 65536);
                        }
                    }

                    sb.WriteShapeProperty("fillToLeft", val1);
                    sb.WriteShapeProperty("fillToRight", val1 + width);
                    sb.WriteShapeProperty("fillToTop", val2);
                    sb.WriteShapeProperty("fillToBottom", val2 + height);
                }
            }

            if (fill.Recolor != null && fill.Recolor.Value)
            {
                sb.WriteShapeProperty("fRecolorFillAsPicture", true); // Default is false
            }

            if (fill.Rotate != null && fill.Rotate.Value)
            {
                sb.WriteShapeProperty("fUseShapeAnchor", true);
            }

            if (fill.AlignShape != null && fill.AlignShape.Value)
            {
                sb.WriteShapeProperty("fillShape", true);
            }

            if (fill.Aspect != null && fill.Aspect.Value == V.ImageAspectValues.AtLeast)
            {
                sb.WriteShapeProperty("fillDztype", "8");
            }
            else if (fill.Aspect != null && fill.Aspect.Value == V.ImageAspectValues.AtMost)
            {
                sb.WriteShapeProperty("fillDztype", "4");
            }

            //if (fill.Size != null)
            //{
            //}
            //if (fill.Source != null)
            //{
            //}
            //if (fill.Position != null)
            //{
            //}
            //if (fill.Origin != null)
            //{
            //}
            //if (fill.Opacity != null)
            //{
            //}
            //if (fill.Opacity2 != null)
            //{
            //}
            //if (fill.On != null)
            //{
            //}

            if (fill.RelationshipId?.Value != null)
            // Textures, pictures and patterns are associated to an embedded image file
            {
                string rId = fill.RelationshipId.Value;
                var rootPart = OpenXmlHelpers.GetRootPart(background);
                var pictWriter = new RtfStringWriter();

                // Get dimensions from the image file
                long imageWidth = 0;
                long imageHeight = 0;
                if (rootPart?.GetPartById(rId) is ImagePart imagePart)
                {
                    MemoryStream stream = new MemoryStream();
                    using (var originalStream = imagePart.GetStream(FileMode.Open, FileAccess.Read))
                    {
                        // In this case we need seeking support,
                        // and the ZipWrappingStream provided by Open XML does not have it.
                        originalStream.CopyTo(stream);
                    }

                    using (stream)
                    {
                        var imageDimensions = ImageHeader.GetDimensions(stream, imagePart.ContentType);
                        imageWidth = imageDimensions.Width;
                        imageHeight = imageDimensions.Height;
                    }
                    if (imageWidth > 0 && imageHeight > 0)
                    {
                        // Wmf is a special case because dimensions are calculated in inches (rather than pixels)
                        if (ImageFormatExtensions.FromMimeType(imagePart.ContentType) == ImageFormat.Wmf)
                        {
                            imageWidth *= 1440;
                            imageHeight *= 1440;
                        }
                        else
                        {
                            // Convert pixels to twips
                            imageWidth = imageWidth * 1440 / 96; // TODO: use the image DPI instead
                            imageHeight = imageHeight * 1440 / 96;
                        }

                        var properties = new PictureProperties()
                        {
                            CropBottom = 0,
                            CropRight = 0,
                            CropLeft = 0,
                            CropTop = 0,
                            WidthGoal = imageWidth,
                            HeightGoal = imageHeight,
                            // Note: picw and pich always seem to be about
                            // picwgoal (or pichgoal) * 1.76 for every background image.
                            // I don't know how this value is calculated.
                            Width = (long)Math.Round(imageWidth * 1.76),
                            Height = (long)Math.Round(imageHeight * 1.76),
                        };
                        ProcessImagePart(rootPart, rId, properties, pictWriter);
                        sb.WriteShapeProperty("fillBlip", pictWriter.ToString());
                    }
                }
            }

            sb.WriteLine("}}}");

            sb.WriteLine(@"\viewbksp1");
        }
    }
}
