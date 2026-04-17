using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class ShadingHelpers
{
    public static bool IsSolid(this Shading shading)
    {
        // Open XML Shading behavior: 
        // - The pure primary color (Fill) is displayed for ShadingPatternValues.Clear
        // or if no pattern (Shading.Val) is specified.
        // - The pure secondary color (Color) is displayed for ShadingPatternValues.Solid. 
        // - Other values (stripes, checkerboard, ...) are displayed as a combination of the two colors.
        //
        // Therefore, this function returns true if: 
        // - the primary color is defined and shading type is Clear or null
        // - OR the secondary color is defined and shading type is Solid
        return (shading.Fill != null && (shading.Val == null || shading.Val == ShadingPatternValues.Clear)) || 
               (shading.Color != null && shading.Val != null && shading.Val == ShadingPatternValues.Solid);
    }
}
