using DocumentFormat.OpenXml;

namespace DocSharp.Renderer
{
    internal static class Emu
    {
        private const double INCH = 72;
        private const double EMU = 914400;

        public static double EmuToPoint(this UInt32Value value)
        {
            var v = System.Convert.ToInt64(value.Value);
            return v.EmuToPoint();
        }

        public static double EmuToPoint(this Int64Value value)
        {
            return value.Value.EmuToPoint();
        }

        public static double EmuToPoint(this long value)
        {
            return value / EMU * INCH;
        }
    }
}
