using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

internal static class RtfShapeTypeMapper
{
    internal static int GetShapeType(ShapeTypeValues shapeType)
    {        
        if (shapeType == ShapeTypeValues.Rectangle)
            return 1;

        else if (shapeType == ShapeTypeValues.RoundRectangle)
            return 2;
        else if (shapeType == ShapeTypeValues.Round1Rectangle)
            return 2;
        else if (shapeType == ShapeTypeValues.Round2SameRectangle)
            return 2;
        else if (shapeType == ShapeTypeValues.Round2DiagonalRectangle)
            return 2;
        else if (shapeType == ShapeTypeValues.SnipRoundRectangle)
            return 2;
        else if (shapeType == ShapeTypeValues.Snip1Rectangle)
            return 2;
        else if (shapeType == ShapeTypeValues.Snip2SameRectangle)
            return 2;
        else if (shapeType == ShapeTypeValues.Snip2DiagonalRectangle)
            return 2;

        else if (shapeType == ShapeTypeValues.Ellipse)
            return 3;
        else if (shapeType == ShapeTypeValues.Diamond)
            return 4;
        else if (shapeType == ShapeTypeValues.Triangle)
            return 5;
        else if (shapeType == ShapeTypeValues.RightTriangle)
            return 6;
        else if (shapeType == ShapeTypeValues.Parallelogram)
            return 7;
        else if (shapeType == ShapeTypeValues.Trapezoid)
            return 8;
        else if (shapeType == ShapeTypeValues.NonIsoscelesTrapezoid)
            return 8;
        else if (shapeType == ShapeTypeValues.Hexagon)
            return 9;
        else if (shapeType == ShapeTypeValues.Octagon)
            return 10;
        else if (shapeType == ShapeTypeValues.Plus)
            return 11;
        else if (shapeType == ShapeTypeValues.MathPlus)
            return 11;
        else if (shapeType == ShapeTypeValues.Star5)
            return 12;
        else if (shapeType == ShapeTypeValues.RightArrow)
            return 13;
        else if (shapeType == ShapeTypeValues.HomePlate)
            return 15;
        else if (shapeType == ShapeTypeValues.Cube)
            return 16;
        else if (shapeType == ShapeTypeValues.Arc)
            return 19;
        else if (shapeType == ShapeTypeValues.Line)
            return 20;
        else if (shapeType == ShapeTypeValues.LineInverse)
            return 20;
        else if (shapeType == ShapeTypeValues.Plaque)
            return 21;
        else if (shapeType == ShapeTypeValues.Can)
            return 22;
        else if (shapeType == ShapeTypeValues.Donut)
            return 23;

        else if (shapeType == ShapeTypeValues.StraightConnector1)
            return 32;
        else if (shapeType == ShapeTypeValues.BentConnector2)
            return 33;
        else if (shapeType == ShapeTypeValues.BentConnector3)
            return 34;
        else if (shapeType == ShapeTypeValues.BentConnector4)
            return 35;
        else if (shapeType == ShapeTypeValues.BentConnector5)
            return 36;
        else if (shapeType == ShapeTypeValues.CurvedConnector2)
            return 37;
        else if (shapeType == ShapeTypeValues.CurvedConnector3)
            return 38;
        else if (shapeType == ShapeTypeValues.CurvedConnector4)
            return 39;
        else if (shapeType == ShapeTypeValues.CurvedConnector5)
            return 40;

        else if (shapeType == ShapeTypeValues.Callout1)
            return 41;
        else if (shapeType == ShapeTypeValues.Callout2)
            return 42;
        else if (shapeType == ShapeTypeValues.Callout3)
            return 43;
        else if (shapeType == ShapeTypeValues.AccentCallout1)
            return 44;
        else if (shapeType == ShapeTypeValues.AccentCallout2)
            return 45;
        else if (shapeType == ShapeTypeValues.AccentCallout3)
            return 46;
        else if (shapeType == ShapeTypeValues.BorderCallout1)
            return 47;
        else if (shapeType == ShapeTypeValues.BorderCallout2)
            return 48;
        else if (shapeType == ShapeTypeValues.BorderCallout3)
            return 49;
        else if (shapeType == ShapeTypeValues.AccentBorderCallout1)
            return 50;
        else if (shapeType == ShapeTypeValues.AccentBorderCallout2)
            return 51;
        else if (shapeType == ShapeTypeValues.AccentBorderCallout3)
            return 52;

        else if (shapeType == ShapeTypeValues.Ribbon)
            return 53;
        else if (shapeType == ShapeTypeValues.Ribbon2)
            return 54;
        else if (shapeType == ShapeTypeValues.Chevron)
            return 55;
        else if (shapeType == ShapeTypeValues.Pentagon)
            return 56;
        else if (shapeType == ShapeTypeValues.NoSmoking)
            return 57;

        else if (shapeType == ShapeTypeValues.Star8)
            return 58;
        else if (shapeType == ShapeTypeValues.Star16)
            return 59;
            // 18 (Seal) is slightly different
        else if (shapeType == ShapeTypeValues.Star32)
            return 60;

        else if (shapeType == ShapeTypeValues.WedgeRectangleCallout)
            return 61;
        else if (shapeType == ShapeTypeValues.WedgeRoundRectangleCallout)
            return 62;
            // 17 (Balloon) is slightly different
        else if (shapeType == ShapeTypeValues.WedgeEllipseCallout)
            return 63;
        else if (shapeType == ShapeTypeValues.Wave)
            return 64;
        else if (shapeType == ShapeTypeValues.FoldedCorner)
            return 65;

        else if (shapeType == ShapeTypeValues.LeftArrow)
            return 66;
        else if (shapeType == ShapeTypeValues.DownArrow)
            return 67;
        else if (shapeType == ShapeTypeValues.UpArrow)
            return 68;
        else if (shapeType == ShapeTypeValues.LeftRightArrow)
            return 69;
        else if (shapeType == ShapeTypeValues.UpDownArrow)
            return 70;

        else if (shapeType == ShapeTypeValues.IrregularSeal1)
            return 71;
        else if (shapeType == ShapeTypeValues.IrregularSeal2)
            return 72;
        else if (shapeType == ShapeTypeValues.LightningBolt)
            return 73;
        else if (shapeType == ShapeTypeValues.Heart)
            return 74;

        else if (shapeType == ShapeTypeValues.QuadArrow)
            return 76;
        else if (shapeType == ShapeTypeValues.LeftArrowCallout)
            return 77;
        else if (shapeType == ShapeTypeValues.RightArrowCallout)
            return 78;
        else if (shapeType == ShapeTypeValues.UpArrowCallout)
            return 79;
        else if (shapeType == ShapeTypeValues.DownArrowCallout)
            return 80;
        else if (shapeType == ShapeTypeValues.LeftRightArrowCallout)
            return 81;
        else if (shapeType == ShapeTypeValues.UpDownArrowCallout)
            return 82;
        else if (shapeType == ShapeTypeValues.QuadArrowCallout)
            return 83;

        else if (shapeType == ShapeTypeValues.Bevel)
            return 84;

        else if (shapeType == ShapeTypeValues.LeftBracket)
            return 85;
        else if (shapeType == ShapeTypeValues.RightBracket)
            return 86;
        else if (shapeType == ShapeTypeValues.LeftBrace)
            return 87;
        else if (shapeType == ShapeTypeValues.RightBrace)
            return 88;

        else if (shapeType == ShapeTypeValues.LeftUpArrow)
            return 89;
        else if (shapeType == ShapeTypeValues.BentUpArrow)
            return 90;
        else if (shapeType == ShapeTypeValues.BentArrow)
            return 91;
        else if (shapeType == ShapeTypeValues.Star24)
            return 92;
        else if (shapeType == ShapeTypeValues.StripedRightArrow)
            return 93;
        else if (shapeType == ShapeTypeValues.NotchedRightArrow)
            return 94;
        else if (shapeType == ShapeTypeValues.BlockArc)
            return 95;
        else if (shapeType == ShapeTypeValues.SmileyFace)
            return 96;

        else if (shapeType == ShapeTypeValues.VerticalScroll)
            return 97;
        else if (shapeType == ShapeTypeValues.HorizontalScroll)
            return 98;

        else if (shapeType == ShapeTypeValues.CircularArrow)
            return 99;
        else if (shapeType == ShapeTypeValues.UTurnArrow)
            return 101;
        else if (shapeType == ShapeTypeValues.CurvedRightArrow)
            return 102;
        else if (shapeType == ShapeTypeValues.CurvedLeftArrow)
            return 103;
        else if (shapeType == ShapeTypeValues.CurvedUpArrow)
            return 104;
        else if (shapeType == ShapeTypeValues.CurvedDownArrow)
            return 105;

        else if (shapeType == ShapeTypeValues.CloudCallout)
            return 106;
        else if (shapeType == ShapeTypeValues.EllipseRibbon)
            return 107;
        else if (shapeType == ShapeTypeValues.EllipseRibbon2)
            return 108;

        else if (shapeType == ShapeTypeValues.FlowChartProcess)
            return 109;
        else if (shapeType == ShapeTypeValues.FlowChartDecision)
            return 110;
        else if (shapeType == ShapeTypeValues.FlowChartInputOutput)
            return 111;
        else if (shapeType == ShapeTypeValues.FlowChartPredefinedProcess)
            return 112;
        else if (shapeType == ShapeTypeValues.FlowChartInternalStorage)
            return 113;
        else if (shapeType == ShapeTypeValues.FlowChartDocument)
            return 114;
        else if (shapeType == ShapeTypeValues.FlowChartMultidocument)
            return 115;
        else if (shapeType == ShapeTypeValues.FlowChartTerminator)
            return 116;
        else if (shapeType == ShapeTypeValues.FlowChartPreparation)
            return 117;
        else if (shapeType == ShapeTypeValues.FlowChartManualInput)
            return 118;
        else if (shapeType == ShapeTypeValues.FlowChartManualOperation)
            return 119;
        else if (shapeType == ShapeTypeValues.FlowChartConnector)
            return 120;
        else if (shapeType == ShapeTypeValues.FlowChartPunchedCard)
            return 121;
        else if (shapeType == ShapeTypeValues.FlowChartPunchedTape)
            return 122;
        else if (shapeType == ShapeTypeValues.FlowChartSummingJunction)
            return 123;
        else if (shapeType == ShapeTypeValues.FlowChartOr)
            return 124;
        else if (shapeType == ShapeTypeValues.FlowChartCollate)
            return 125;
        else if (shapeType == ShapeTypeValues.FlowChartSort)
            return 126;
        else if (shapeType == ShapeTypeValues.FlowChartExtract)
            return 127;
        else if (shapeType == ShapeTypeValues.FlowChartMerge)
            return 128;
        else if (shapeType == ShapeTypeValues.FlowChartOfflineStorage)
            return 129;
        else if (shapeType == ShapeTypeValues.FlowChartOnlineStorage)
            return 130;
        else if (shapeType == ShapeTypeValues.FlowChartMagneticTape)
            return 131;
        else if (shapeType == ShapeTypeValues.FlowChartMagneticDisk)
            return 132;
        else if (shapeType == ShapeTypeValues.FlowChartMagneticDrum)
            return 133;
        else if (shapeType == ShapeTypeValues.FlowChartDisplay)
            return 134;
        else if (shapeType == ShapeTypeValues.FlowChartDelay)
            return 135;
        else if (shapeType == ShapeTypeValues.FlowChartAlternateProcess)
            return 176;
        else if (shapeType == ShapeTypeValues.FlowChartOffpageConnector)
            return 177;

        else if (shapeType == ShapeTypeValues.LeftRightUpArrow)
            return 182;
        else if (shapeType == ShapeTypeValues.Sun)
            return 183;
        else if (shapeType == ShapeTypeValues.Moon)
            return 184;
        else if (shapeType == ShapeTypeValues.BracketPair)
            return 185;
        else if (shapeType == ShapeTypeValues.BracePair)
            return 186;
        else if (shapeType == ShapeTypeValues.Star4)
            return 187;
        else if (shapeType == ShapeTypeValues.DoubleWave)
            return 188;

        else if (shapeType == ShapeTypeValues.ActionButtonBlank)
            return 189;
        else if (shapeType == ShapeTypeValues.ActionButtonHome)
            return 190;
        else if (shapeType == ShapeTypeValues.ActionButtonHelp)
            return 191;
        else if (shapeType == ShapeTypeValues.ActionButtonInformation)
            return 192;
        else if (shapeType == ShapeTypeValues.ActionButtonForwardNext)
            return 193;
        else if (shapeType == ShapeTypeValues.ActionButtonBackPrevious)
            return 194;
        else if (shapeType == ShapeTypeValues.ActionButtonEnd)
            return 195;
        else if (shapeType == ShapeTypeValues.ActionButtonBeginning)
            return 196;
        else if (shapeType == ShapeTypeValues.ActionButtonReturn)
            return 197;
        else if (shapeType == ShapeTypeValues.ActionButtonDocument)
            return 198;
        else if (shapeType == ShapeTypeValues.ActionButtonSound)
            return 199;
        else if (shapeType == ShapeTypeValues.ActionButtonMovie)
            return 200;

        // TODO:
        // - for NonIsoscelesTrapezoid, LineInverse, Round1Rectangle, Round2SameRectangle,
        // Round2DiagonalRectangle, SnipRoundRectangle, Snip1Rectangle, Snip2SameRectangle,
        // Snip2DiagonalRectangle, MathPlus we should set other properties to match the shape appearance
        // - for other shapes (Heptagon, Decagon, Dodecagon, Star6, Star7, Star10, Star12,
        // Teardrop, PieWedge, Pie, LeftCircularArrow, LeftRightCircularArrow, SwooshArrow, 
        // Frame, HalfFrame, Corner, DiagonalStripe, Chord, Cloud, LeftRightRibbon, 
        // Gear6, Gear9, Funnel, MathPlus, MathMinus, MathMultiply, MathDivide, MathEqual, MathNotEqual, 
        // CornerTabs, SquareTabs, PlaqueTabs, 
        // ChartX, ChartStar, ChartPlus), we should return 0 (custom / not AutoShape) and set other properties. 
        //
        // When using Word Automation for testing, essentially shape types 140-183
        // are not directly supported in RTF
        return 1;
    }
}
