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
}
