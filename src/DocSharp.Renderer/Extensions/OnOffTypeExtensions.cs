using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer
{
    internal static class OnOffTypeExtensions
    {
        public static bool IsOn(this OnOffType onOffType, bool ifOnOffTypeNull = false, bool ifOnOffValueNull = false)
        {
            return onOffType?.Val.IsOn(ifOnOffValueNull) ?? ifOnOffTypeNull;
        }

        public static bool IsOn(this OnOffValue onOffValue, bool ifNull = false)
        {
            return onOffValue?.Value ?? ifNull;
        }
    }
}
