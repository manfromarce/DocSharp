using System;

namespace DocSharp.Helpers;

public static class MathHelpers
{
    public static float? Negate(float? val)
    {
        if (val == null)
            return null;
        else 
            return -val.Value;
    }

    // Math.Clamp is not available in .NET Framework
    public static long Clamp(long value, long min, long max)
    {
        return Math.Min(max, Math.Max(min, value));
    }

    public static ulong Clamp(ulong value, ulong min, ulong max)
    {
        return Math.Min(max, Math.Max(min, value));
    }

    public static decimal Clamp(decimal value, decimal min, decimal max)
    {
        return Math.Min(max, Math.Max(min, value));
    }

    public static double Clamp(double value, double min, double max)
    {
        return Math.Min(max, Math.Max(min, value));
    }
}
