using System;
using DocumentFormat.OpenXml;

namespace DocSharp.Renderer
{
    internal static class EightPoint
    {
        private const float Factor = 8;

        public static float EpToPoint(this StringValue value, float ifNull = 0)
        {
            if (value?.Value == null)
            {
                return ifNull;
            }

            var v = Convert.ToInt32(value.Value);
            return v.EpToPoint();
        }

        public static float EpToPoint(this UInt32Value value, float ifNull = 0)
        {
            if (value?.Value == null)
            {
                return ifNull;
            }

            var v = Convert.ToInt32(value.Value);
            return v.EpToPoint();
        }

        public static float EpToPoint(this int value)
            => value.ToFloat(Factor);
    }
}
