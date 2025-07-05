using System;
using System.IO;
using System.Text;
using DocSharp.Helpers;

namespace DocSharp.Writers;

public abstract class BaseStringWriter : IDisposable
{
    protected StringBuilder sb = new StringBuilder();
    
    public TextWriter? ExternalWriter;
    public StringBuilder StringBuilder => sb;

    public string NewLine { get; set; } = Environment.NewLine;

    public void Dispose()
    {
        sb.Clear();
        ExternalWriter?.Flush();
    }

    public override string ToString()
    {
        return sb.ToString();
    }

    public virtual bool EndsWithNewLine()
    {
        return sb.EndsWithNewLine();
    }

    public virtual void Append(string text)
    {
        ExternalWriter?.Write(text);
        sb.Append(text);
    }

    public virtual void Append(char c)
    {
        ExternalWriter?.Write(c);
        sb.Append(c);
    }

    public virtual void AppendLine()
    {
        ExternalWriter?.Write(NewLine);
        sb.Append(NewLine);
    }

    public virtual void AppendLine(string text)
    {
        Append(text);
        AppendLine();
    }

    public virtual void AppendLine(char c)
    {
        Append(c);
        AppendLine();
    }

    public virtual void Append(int value)
    {
        Append(value.ToStringInvariant());
    }

    public virtual void Append(double value)
    {
        Append(value.ToStringInvariant());
    }

    public virtual void Append(float value)
    {
        Append(value.ToStringInvariant());
    }

    public virtual void Append(decimal value)
    {
        Append(value.ToStringInvariant());
    }

    public virtual void Append(long value)
    {
        Append(value.ToStringInvariant());
    }

    public virtual void Append(short value)
    {
        Append(value.ToStringInvariant());
    }

    public virtual void Append(byte value)
    {
        Append(value.ToStringInvariant());
    }

    public virtual void Append(uint value)
    {
        Append(value.ToStringInvariant());
    }

    public virtual void Append(ulong value)
    {
        Append(value.ToStringInvariant());
    }

    public virtual void Append(ushort value)
    {
        Append(value.ToStringInvariant());
    }

    public void AppendFormat(string format, params object?[] args)
    {
        sb.AppendFormat(format, args);
        ExternalWriter?.Write(string.Format(format, args));
    }

    public void AppendFormat(IFormatProvider? provider, string format, params object?[] args)
    {
        sb.AppendFormat(provider, format, args);
        ExternalWriter?.Write(string.Format(provider, format, args));
    }
}
