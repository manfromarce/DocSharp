using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
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
        Flush();
    }

    public override string ToString()
    {
        return sb.ToString();
    }

    public virtual void EnsureEmptyLine()
    {
        if (sb.Length == 0)
            return; // Don't add space if we are just at the start of the document

        if (!EndsWithParagraph()) // If the string ends with 1 or 0 new line chars
        {
            WriteLine(); 
        }
        if (!EndsWithParagraph()) // If the string ends with only 1 new line char
        {
            WriteLine();
        }
    }

    public virtual bool EndsWithNewLine()
    {
        return sb.EndsWithNewLine();
    }

    public virtual bool EndsWithParagraph()
    {
        return sb.EndsWithEmptyLine();
    }

    public virtual void Write(string? text)
    {
        if (text != null)
        {
            ExternalWriter?.Write(text);
            sb.Append(text);
        }
    }

    public virtual void Write(char c)
    {
        ExternalWriter?.Write(c);
        sb.Append(c);
    }

    public virtual void Write(char[] c)
    {
        ExternalWriter?.Write(c);
        sb.Append(c);
    }

    public virtual void Write(char[] buffer, int index, int count)
    {
        sb.Append(buffer, index, count);
        ExternalWriter?.Write(buffer, index, count);
    }

    public virtual void WriteLine()
    {
        ExternalWriter?.Write(NewLine);
        sb.Append(NewLine);
    }

    public virtual void WriteLine(string text)
    {
        Write(text);
        WriteLine();
    }

    public virtual void WriteLine(char c)
    {
        Write(c);
        WriteLine();
    }

    public virtual void WriteLine(char[] c)
    {
        Write(c);
        WriteLine();
    }

    public virtual void Write(int value)
    {
        Write(value.ToStringInvariant());
    }

    public virtual void Write(double value)
    {
        Write(value.ToStringInvariant());
    }

    public virtual void Write(float value)
    {
        Write(value.ToStringInvariant());
    }

    public virtual void Write(decimal value)
    {
        Write(value.ToStringInvariant());
    }

    public virtual void Write(long value)
    {
        Write(value.ToStringInvariant());
    }

    public virtual void Write(short value)
    {
        Write(value.ToStringInvariant());
    }

    public virtual void Write(byte value)
    {
        Write(value.ToStringInvariant());
    }

    public virtual void Write(uint value)
    {
        Write(value.ToStringInvariant());
    }

    public virtual void Write(ulong value)
    {
        Write(value.ToStringInvariant());
    }

    public virtual void Write(ushort value)
    {
        Write(value.ToStringInvariant());
    }

    public void WriteFormat(string format, params object?[] args)
    {
        sb.AppendFormat(format, args);
        ExternalWriter?.Write(string.Format(format, args));
    }

    public void WriteFormat(IFormatProvider? provider, string format, params object?[] args)
    {
        sb.AppendFormat(provider, format, args);
        ExternalWriter?.Write(string.Format(provider, format, args));
    }

    public void Flush()
    {
        ExternalWriter?.Flush();
    }

    public Task? FlushAsync()
    {
        return ExternalWriter?.FlushAsync();
    }
}
