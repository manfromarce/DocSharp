using System;

namespace DocSharp.Wmf2Svg.Wmf;

public sealed class WmfParseException : Exception
{
    public WmfParseException()
    {
    }

    public WmfParseException(string message) : base(message)
    {
    }

    public WmfParseException(string message, Exception innerException) : base(message, innerException)
    {
    }
}