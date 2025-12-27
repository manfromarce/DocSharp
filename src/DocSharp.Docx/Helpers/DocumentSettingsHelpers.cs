using System;
using System.Globalization;

namespace DocSharp.Docx;

public static class DocumentSettingsHelpers
{    
    /// <summary>
    /// Gets the default page width in twips. The default page size is A4 (21 x 29.7 cm) for regions using metric units and Letter for regions using imperial units.
    /// </summary>
    /// <returns></returns>
    public static int GetDefaultPageWidth()
    {
        return RegionInfo.CurrentRegion.IsMetric ? 11906 : 12240;
    }

    /// <summary>
    /// Gets the default page width in twips. The default page size is A4 (21 x 29.7 cm) for regions using metric units and Letter for regions using imperial units.
    /// </summary>
    /// <returns></returns>
    public static int GetDefaultPageHeight()
    {
        return RegionInfo.CurrentRegion.IsMetric ? 16838 : 15839;
    }

    /// <summary>
    /// Gets the default page left margin (2 cm) in twips.
    /// </summary>
    /// <returns></returns>
    public static int GetDefaultPageLeftMargin()
    {
        return 1134;
    }

    /// <summary>
    /// Gets the default page top margin (2.5 cm) in twips.
    /// </summary>
    /// <returns></returns>
    public static int GetDefaultPageTopMargin()
    {
        return 1417;
    }

    /// <summary>
    /// Gets the default page right margin (2 cm) in twips.
    /// </summary>
    /// <returns></returns>
    public static int GetDefaultPageRightMargin()
    {
        return 1134;
    }

    /// <summary>
    /// Gets the default page bottom margin (2 cm) in twips.
    /// </summary>
    /// <returns></returns>
    public static int GetDefaultPageBottomMargin()
    {
        return 1134;
    }
}