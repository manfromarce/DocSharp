/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System;
using System.Globalization;

namespace DocSharp;

/// <summary>
/// Represents a Html Unit (ie: 120px, 10em, ...).
/// </summary>
[System.Diagnostics.DebuggerDisplay("Unit: {Value} {Type}")]
readonly struct Unit
{
    /// <summary>Represents an empty unit (not defined).</summary>
    public static readonly Unit Empty = new Unit();
    /// <summary>Represents an Auto unit.</summary>
    public static readonly Unit Auto = new Unit(UnitMetric.Auto, 0L);

    private readonly UnitMetric type;
    private readonly double value;
    private readonly long valueInEmus;

    public Unit(UnitMetric type, double value)
    {
        this.type = type;
        this.value = value;
        this.valueInEmus = CalculateEMUs(type, value);
    }

    public static Unit Parse(string? str, UnitMetric defaultMetric = UnitMetric.Unitless)
    {
        if (str == null) return Unit.Empty;

        str = str.Trim().ToLowerInvariant();
        int length = str.Length;
        int digitLength = -1;
        for (int i = 0; i < length; i++)
        {
            char ch = str[i];
            if ((ch < '0' || ch > '9') && ch != '-' && ch != '.' && ch != ',')
                break;

            digitLength = i;
        }
        if (digitLength == -1)
        {
            // No digits in the width, we ignore this style
            return str == "auto"? Unit.Auto : Unit.Empty;
        }

        UnitMetric type;
        if (digitLength < length - 1)
            type = UnitMetricHelper.ToUnitMetric(str.Substring(digitLength + 1).Trim());
        else
            type = defaultMetric;

        string v = str.Substring(0, digitLength + 1);
        double value;
        try
        {
            value = Convert.ToDouble(v, CultureInfo.InvariantCulture);

            if (value < short.MinValue || value > short.MaxValue)
                return Unit.Empty;
        }
        catch (FormatException)
        {
            return Unit.Empty;
        }
        catch (ArithmeticException)
        {
            return Unit.Empty;
        }

        return new Unit(type, value);
    }

    /// <summary>
    /// Gets the value expressed in English Metrics Units.
    /// </summary>
    internal static long CalculateEMUs(UnitMetric type, double value)
    {
        /* Compute width and height in English Metrics Units.
            * There are 360000 EMUs per centimeter, 914400 EMUs per inch, 12700 EMUs per point
            * widthInEmus = widthInPixels / HorizontalResolutionInDPI * 914400
            * heightInEmus = heightInPixels / VerticalResolutionInDPI * 914400
            * 
            * According to 1 px ~= 9525 EMU -> 914400 EMU per inch / 9525 EMU = 96 dpi
            * So Word use 96 DPI printing which seems fair.
            * http://hastobe.net/blogs/stevemorgan/archive/2008/09/15/howto-insert-an-image-into-a-word-document-and-display-it-using-openxml.aspx
            * http://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
            *
            * The list of units supported are explained here: http://www.w3schools.com/css/css_units.asp
            */

        switch (type)
        {
            case UnitMetric.Auto:
            case UnitMetric.Unitless:
            case UnitMetric.Percent: return 0L; // not applicable
            case UnitMetric.Emus: return (long) value;
            case UnitMetric.Inch: return (long) (value * 914400L);
            case UnitMetric.Centimeter: return (long) (value * 360000L);
            case UnitMetric.Point: return (long) (value * 12700L); // 1 point = 1/72 inch

            case UnitMetric.HundrethsOfInch: return (long) (value * 9144L);
            case UnitMetric.Twip: return (long) (value * 635L); // 1 twip = 1/20 point = 12700/20 EMUs = 635 EMUs
            case UnitMetric.Pica: return (long) (value * 152400L); // 1 pica = 1/6 inch = 12 pt = 152400 EMUs

            case UnitMetric.Millimeter: return (long) (value * 36000L);
            case UnitMetric.Himetric: return (long) (value * 360L); // 1 himetric = 1/100 mm

            case UnitMetric.EM:
                // Considering 1em = 12pt (http://sureshjain.wordpress.com/2007/07/06/53/)    
                return (long) (value * 152400);
            case UnitMetric.Ex: // Considering half of EM
                return (long) (value * 152400) / 2;

            case UnitMetric.Diu:
                // 1 DIU = 1/96 inch = 914400/96 EMUs = 9525 EMUs
                return (long) (value * 9525L);
            case UnitMetric.Pixel:
                // Considering 96 DPI as Microsoft Word uses this value
                return (long) (value * 9525L);
            default: goto case UnitMetric.Pixel;
        }
    }

    /// <summary>
    /// Gets the value expressed in twips / DXA.
    /// </summary>
    internal static long CalculateTwips(UnitMetric type, double value)
    {
        switch (type)
        {
            case UnitMetric.Auto:
            case UnitMetric.Unitless:
            case UnitMetric.Percent: return 0L; // not applicable

            case UnitMetric.Twip: return (long)value;
            case UnitMetric.Point: return (long)(value * 20); // 1 twip = 1/20 point
            case UnitMetric.Inch: return (long)(value * 1440); // 1 twip = 1/1440 inch
            case UnitMetric.HundrethsOfInch: return (long)(value * 14.4); // 1/100 inch
            case UnitMetric.Pica: return (long)(value * 240); // 1 pica = 1/6 inch

            case UnitMetric.Centimeter: return (long)((value * 1440) / 2.54); // 1 cm = 1/2.54 inch
            case UnitMetric.Millimeter: return (long)((value * 1440) / 25.4); // 1 mm = 1/25.4 inch
            case UnitMetric.Himetric: return (long)((value * 14.4) / 25.4); // 1 himetric = 1/100 mm; twips = mm*1440/25.4 --> twips = (1440/25.4)/100 himetric

            case UnitMetric.Emus: return (long)(value / 635); // 1 inch = 914400 EMUs = 1440 twips --> 1 twip = 914400 / 1440 = 635

            case UnitMetric.EM: return (long)(value * 20 * 12); // Considering 1 em = 12 pt
            case UnitMetric.Ex: return (long)(value * 20 * 12 / 2); // Considering half of em
            
            case UnitMetric.Diu: return (long)(value * 15); // 1 DIU = 1/96 inch
            case UnitMetric.Pixel: return (long)(value * 15); // Considering 96 DPI

            default: goto case UnitMetric.Pixel;
        }
    }

    //____________________________________________________________________
    //

    /// <summary>
    /// Gets the type of unit (pixel, percent, point, ...)
    /// </summary>
    public UnitMetric Type
    {
        get { return type; }
    }

    /// <summary>
    /// Gets the value of this unit.
    /// </summary>
    public double Value
    {
        get { return value; }
    }

    /// <summary>
    /// Gets the value expressed in English Metrics Unit.
    /// </summary>
    public long ValueInEmus
    {
        get { return valueInEmus; }
    }

    /// <summary>
    /// Gets the value expressed in Dxa unit.
    /// </summary>
    public long ValueInDxa
    {
        get { return (long) (((double) valueInEmus / 914400L) * 20 * 72); }
    }

    /// <summary>
    /// Gets the value expressed in Pixel unit.
    /// </summary>
    public int ValueInPx
    {
        get { return (int) (type == UnitMetric.Pixel ? this.value : (float) valueInEmus / 914400L * 96); }
    }

    /// <summary>
    /// Gets the value expressed in Point unit.
    /// </summary>
    public double ValueInPoint
    {
        get { return (double) (type == UnitMetric.Point ? this.value : (float) valueInEmus / 12700L); }
    }

    /// <summary>
    /// Gets the value expressed in 1/8 of a Point
    /// IMPORTANT: Use this for borders, as OpenXML expresses Border Width in 1/8 of points,
    /// with a minimum value of 2 (1/4 of a point) and a maximum value of 96 (12 points).
    /// </summary>
    public double ValueInEighthPoint
    {
        get { return ValueInPoint * 8; }
    }

    /// <summary>
    /// Gets whether the unit is well formed and not empty.
    /// </summary>
    public bool IsValid
    {
        get { return this.Type != UnitMetric.Unknown; }
    }

    /// <summary>
    /// Gets whether the unit is well formed and not absolute nor auto.
    /// </summary>
    public bool IsFixed
    {
        get { return IsValid && Type != UnitMetric.Percent && Type != UnitMetric.Auto; }
    }
}
