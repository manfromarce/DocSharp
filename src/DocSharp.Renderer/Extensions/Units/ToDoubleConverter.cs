using DocumentFormat.OpenXml;

namespace DocSharp.Renderer
{
    internal static class ToDoubleConverter
    {
        public static double ToDouble(this Int32Value value, double factor)
        {
            if (value == null || !value.HasValue)
            {
                return 0;
            }

            return value.Value.ToDouble(factor);
        }

        public static double ToDouble(this UInt32Value value, double factor)
        {
            if (value == null || !value.HasValue)
            {
                return 0;
            }

            return value.Value.ToDouble(factor);
        }

        public static double ToDouble(this uint value, double factor)
        {
            return value / factor;
        }

        public static double ToDouble(this int value, double factor)
        {
            return value / factor;
        }

        public static double ToDouble(this double value, double factor)
        {
            return value / factor;
        }
    }
}
