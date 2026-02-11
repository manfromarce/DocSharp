using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Helpers;

public static class SafeTypeConverter
{
    // Implicit conversions: 
    // - any integer type --> double / float / decimal
    // - float --> double
    // - signed integer type --> bigger signed integer type
    // - unsigned integer type --> bigger signed or unsigned integer type

    public static long ToLong(this ulong value)
    {
        ushort x = 3;
        int k = x;
        return (long)Math.Min(value, long.MaxValue);
    }

    public static uint ToUint(this ulong value)
    {
        return (uint)Math.Min(value, uint.MaxValue);
    }

    public static int ToInt(this ulong value)
    {
        return (int)Math.Min(value, int.MaxValue);
    }

    public static ushort ToUshort(this ulong value)
    {
        return (ushort)Math.Min(value, ushort.MaxValue);
    }

    public static short ToShort(this ulong value)
    {
        return (short)Math.Min(value, (ulong)short.MaxValue);
    }

    public static byte ToByte(this ulong value)
    {
        return (byte)Math.Min(value, byte.MaxValue);        
    }

    public static sbyte ToSbyte(this ulong value)
    {
        return (sbyte)Math.Min(value, (ulong)sbyte.MaxValue);
    }

    public static ulong ToUlong(this long value)
    {
        return (ulong)Math.Abs(value);
    }

    public static uint ToUint(this long value)
    {
        return (uint)Math.Min(Math.Abs(value), uint.MaxValue);
    }

    public static int ToInt(this long value)
    {
        return (int)MathHelpers.Clamp(value, int.MinValue, int.MaxValue);
    }

    public static ushort ToUshort(this long value)
    {
        return (ushort)Math.Min(Math.Abs(value), ushort.MaxValue);
    }

    public static short ToShort(this long value)
    {
        return (short)MathHelpers.Clamp(value, short.MinValue, short.MaxValue);
    }

    public static byte ToByte(this long value)
    {
        return (byte)Math.Min(Math.Abs(value), byte.MaxValue);
    }

    public static sbyte ToSbyte(this long value)
    {
        return (sbyte)MathHelpers.Clamp(value, sbyte.MinValue, sbyte.MaxValue);
    }

    public static ulong ToUlong(this decimal value)
    {
        return (ulong)Math.Min(Math.Abs(Math.Round(value)), ulong.MaxValue);
    }

    public static long ToLong(this decimal value)
    {
        return (long)MathHelpers.Clamp(Math.Round(value), long.MinValue, long.MaxValue);
    }

    public static uint ToUint(this decimal value)
    {
        return (uint)Math.Min(Math.Abs(Math.Round(value)), uint.MaxValue);        
    }

    public static int ToInt(this decimal value)
    {
        return (int)MathHelpers.Clamp(Math.Round(value), int.MinValue, int.MaxValue);
    }

    public static ushort ToUshort(this decimal value)
    {
        return (ushort)Math.Min(Math.Abs(Math.Round(value)), ushort.MaxValue);        
    }

    public static short ToShort(this decimal value)
    {
        return (short)MathHelpers.Clamp(Math.Round(value), short.MinValue, short.MaxValue);
    }

    public static byte ToByte(this decimal value)
    {
        return (byte)Math.Min(Math.Abs(Math.Round(value)), byte.MaxValue);        
    }

    public static sbyte ToSbyte(this decimal value)
    {
        return (sbyte)MathHelpers.Clamp(Math.Round(value), sbyte.MinValue, sbyte.MaxValue);
    }

    public static double ToDouble(this decimal value)
    {
        return decimal.ToDouble(value);
        // return decimal.ToSingle(Math.Round(value, 15));
    }

    public static float ToFloat(this decimal value)
    {
        return decimal.ToSingle(value);
        // return decimal.ToSingle(Math.Round(value, 6));
    }

    public static ulong ToUlong(this double value)
    {
        return (ulong)Math.Min(Math.Abs(Math.Round(value)), ulong.MaxValue);
    }

    public static long ToLong(this double value)
    {
        return (long)MathHelpers.Clamp(Math.Round(value), long.MinValue, long.MaxValue);
    }

    public static uint ToUint(this double value)
    {
        return (uint)Math.Min(Math.Abs(Math.Round(value)), uint.MaxValue);
    }

    public static int ToInt(this double value)
    {
        return (int)MathHelpers.Clamp(Math.Round(value), int.MinValue, int.MaxValue);
    }

    public static ushort ToUshort(this double value)
    {
        return (ushort)Math.Min(Math.Abs(Math.Round(value)), ushort.MaxValue);
    }

    public static short ToShort(this double value)
    {
        return (short)MathHelpers.Clamp(Math.Round(value), short.MinValue, short.MaxValue);
    }

    public static byte ToByte(this double value)
    {
        return (byte)Math.Min(Math.Abs(Math.Round(value)), byte.MaxValue);        
    }
    
    public static sbyte ToSbyte(this double value)
    {
        return (sbyte)MathHelpers.Clamp(Math.Round(value), sbyte.MinValue, sbyte.MaxValue);
    }

    public static decimal ToDecimal(this double value)
    {        
        if (double.IsNaN(value)) return decimal.Zero;

        if (value >= decimal.MaxValue.ToDouble()) return decimal.MaxValue;
        if (value <= decimal.MinValue.ToDouble()) return decimal.MinValue;
    
        return (decimal)value;
    }

    public static float ToFloat(this double value)
    {
        return (float)Math.Round(value, 6);
    }

    public static ulong ToUlong(this float value)
    {
        return (ulong)Math.Min(Math.Abs(Math.Round(value)), ulong.MaxValue);
    }

    public static long ToLong(this float value)
    {
        return (long)MathHelpers.Clamp(Math.Round(value), long.MinValue, long.MaxValue);
    }

    public static uint ToUint(this float value)
    {
        return (uint)Math.Min(Math.Abs(Math.Round(value)), uint.MaxValue);
    }

    public static int ToInt(this float value)
    {
        return (int)MathHelpers.Clamp(Math.Round(value), int.MinValue, int.MaxValue);
    }

    public static ushort ToUshort(this float value)
    {
        return (ushort)Math.Min(Math.Abs(Math.Round(value)), ushort.MaxValue);
    }

    public static short ToShort(this float value)
    {
        return (short)MathHelpers.Clamp(Math.Round(value), short.MinValue, short.MaxValue);
    }

    public static byte ToByte(this float value)
    {
        return (byte)Math.Min(Math.Abs(Math.Round(value)), byte.MaxValue);        
    }
    
    public static sbyte ToSbyte(this float value)
    {
        return (sbyte)MathHelpers.Clamp(Math.Round(value), sbyte.MinValue, sbyte.MaxValue);
    }

    public static decimal ToDecimal(this float value)
    {        
        if (float.IsNaN(value)) return decimal.Zero;

        if (value >= decimal.MaxValue.ToFloat()) return decimal.MaxValue;
        if (value <= decimal.MinValue.ToFloat()) return decimal.MinValue;
    
        return (decimal)value;
    }

    public static int ToInt(this uint value)
    {
        return (int)Math.Min(value, int.MaxValue);
    }

    public static ushort ToUshort(this uint value)
    {
        return (ushort)Math.Min(value, ushort.MaxValue);
    }

    public static short ToShort(this uint value)
    {
        return (short)Math.Min(value, short.MaxValue);
    }

    public static byte ToByte(this uint value)
    {
        return (byte)Math.Min(value, byte.MaxValue);        
    }

    public static sbyte ToSbyte(this uint value)
    {
        return (sbyte)Math.Min(value, sbyte.MaxValue);
    }

    public static ulong ToUlong(this int value)
    {
        return (ulong)Math.Abs(value);
    }

    public static uint ToUint(this int value)
    {
        return (uint)Math.Abs(value);
    }

    public static ushort ToUshort(this int value)
    {
        return (ushort)Math.Min(Math.Abs(value), ushort.MaxValue);
    }

    public static short ToShort(this int value)
    {
        return (short)MathHelpers.Clamp(value, short.MinValue, short.MaxValue);
    }

    public static byte ToByte(this int value)
    {
        return (byte)Math.Min(Math.Abs(value), byte.MaxValue);
    }

    public static sbyte ToSbyte(this int value)
    {
        return (sbyte)MathHelpers.Clamp(value, sbyte.MinValue, sbyte.MaxValue);
    }

    public static short ToShort(this ushort value)
    {
        return (short)Math.Min(value, short.MaxValue);
    }

    public static byte ToByte(this ushort value)
    {
        return (byte)Math.Min(value, byte.MaxValue);        
    }

    public static sbyte ToSbyte(this ushort value)
    {
        return (sbyte)Math.Min(value, sbyte.MaxValue);
    }

    public static ulong ToUlong(this short value)
    {
        return (ulong)Math.Abs(value);
    }

    public static uint ToUint(this short value)
    {
        return (uint)Math.Abs(value);
    }

    public static ushort ToUshort(this short value)
    {
        return (ushort)Math.Abs(value);
    }

    public static byte ToByte(this short value)
    {
        return (byte)Math.Min(Math.Abs(value), byte.MaxValue);
    }

    public static sbyte ToSbyte(this short value)
    {
        return (sbyte)MathHelpers.Clamp(value, sbyte.MinValue, sbyte.MaxValue);
    }

    public static sbyte ToSbyte(this byte value)
    {
        return (sbyte)Math.Min(value, sbyte.MaxValue);
    }

    public static ulong ToUlong(this sbyte value)
    {
        return (ulong)Math.Abs(value);
    }

    public static uint ToUint(this sbyte value)
    {
        return (uint)Math.Abs(value);
    }

    public static ushort ToUshort(this sbyte value)
    {
        return (ushort)Math.Abs(value);
    }

    public static byte ToByte(this sbyte value)
    {
        return (byte)Math.Abs(value);
    }
}