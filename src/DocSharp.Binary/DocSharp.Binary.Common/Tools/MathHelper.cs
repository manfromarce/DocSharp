using System;
using System.Globalization;

namespace DocSharp.Binary.Tools
{
    public static class MathHelper
    {
        /// <summary>
        /// Alternative to Math.Clamp for .NET Framework / .NET Standard compatibility
        /// </summary>
        /// <param name="val"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns></returns>
        public static int Clamp(int val, int min, int max)
        {
            return (val < min) ? min : (val > max) ? max : val;
        }

        /// <summary>
        /// Alternative to Math.Clamp for .NET Framework / .NET Standard compatibility
        /// </summary>
        /// <param name="val"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns></returns>
        /// <remarks>Overload for double values</remarks>
        public static double Clamp(double val, double min, double max)
        {
            return (val < min) ? min : (val > max) ? max : val;
        }

        /// <summary>
        /// Alternative to Math.Clamp for .NET Framework / .NET Standard compatibility
        /// </summary>
        /// <param name="val"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns></returns>
        /// <remarks>Overload for float values</remarks>
        public static float Clamp(float val, float min, float max)
        {
            return (val < min) ? min : (val > max) ? max : val;
        }

        /// <summary>
        /// Alternative to Math.Clamp for .NET Framework / .NET Standard compatibility
        /// </summary>
        /// <param name="val"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns></returns>
        /// <remarks>Overload for decimal values</remarks>
        public static decimal Clamp(decimal val, decimal min, decimal max)
        {
            return (val < min) ? min : (val > max) ? max : val;
        }

        /// <summary>
        /// Alternative to Math.Clamp for .NET Framework / .NET Standard compatibility
        /// </summary>
        /// <param name="val"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns></returns>
        /// <remarks>Overload for uint values</remarks>
        public static uint Clamp(uint val, uint min, uint max)
        {
            return (val < min) ? min : (val > max) ? max : val;
        }

        /// <summary>
        /// Alternative to Math.Clamp for .NET Framework / .NET Standard compatibility
        /// </summary>
        /// <param name="val"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns></returns>
        /// <remarks>Overload for long values</remarks>
        public static long Clamp(long val, long min, long max)
        {
            return (val < min) ? min : (val > max) ? max : val;
        }

        /// <summary>
        /// Alternative to Math.Clamp for .NET Framework / .NET Standard compatibility
        /// </summary>
        /// <param name="val"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns></returns>
        /// <remarks>Overload for ulong values</remarks>
        public static ulong Clamp(ulong val, ulong min, ulong max)
        {
            return (val < min) ? min : (val > max) ? max : val;
        }        
    }
}