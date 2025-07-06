using DocumentFormat.OpenXml;

namespace DocSharp.Renderer
{
    internal static class TwentiethPoint
    {
        private const double Factor = 20;

        public static double DxaToPoint(this Int32Value value)
            => value.ToDouble(Factor);

        public static double DxaToPoint(this UInt32Value value)
            => value.ToDouble(Factor);

        public static double DxaToPoint(this uint value)
            => value.ToDouble(Factor);

        public static double DxaToPoint(this int value)
            => value.ToDouble(Factor);

        public static double DxaToPoint(this double value)
            => value.ToDouble(Factor);
    }
}
