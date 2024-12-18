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

namespace DocSharp;

/// <summary>
/// Specifies the measurement values of a Html Unit.
/// </summary>
public enum UnitMetric
{
    Unknown,
    Percent,
    Inch,
    Centimeter,
    Millimeter,
    /// <summary>1em is equal to the current font size.</summary>
    EM,
    /// <summary>one ex is the x-height of a font (x-height is usually about half the font-size)</summary>
    Ex,
    Point,
    Pica,
    Pixel,

    // this value is not parsed but can be used internally
    Emus,

    /// <summary>Not convertible to any other units.</summary>
    Auto,
    /// <summary>Raw value, not convertible to any other units</summary>
    Unitless
}

public static class UnitMetricHelper
{
    /// <summary>
    /// Converts value of the specified UnitMetric to EMUs.
    /// </summary>
    public static long ConvertToEmus(double value, UnitMetric unitType)
    {
        return Unit.ComputeInEmus(unitType, value);
    }

    internal static UnitMetric ToUnitMetric(string? type)
    {
        if (type == null) return UnitMetric.Unitless;
        return type.ToLowerInvariant() switch
        {
            "%" => UnitMetric.Percent,
            "in" => UnitMetric.Inch,
            "cm" => UnitMetric.Centimeter,
            "mm" => UnitMetric.Millimeter,
            "em" => UnitMetric.EM,
            "ex" => UnitMetric.Ex,
            "pt" => UnitMetric.Point,
            "pc" => UnitMetric.Pica,
            "px" => UnitMetric.Pixel,
            _ => UnitMetric.Unknown,
        };
    }
}

