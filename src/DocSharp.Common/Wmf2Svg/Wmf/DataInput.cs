using System;
using System.IO;

namespace DocSharp.Wmf2Svg.Wmf;

internal sealed class DataInput : IDisposable
{
    private readonly Stream _stream;
    private readonly bool _isLittleEndian;
    private readonly byte[] _buf = new byte[4];
    private int _count;

    public DataInput(Stream stream, bool isLittleEndian)
    {
        _stream = stream;
        _isLittleEndian = isLittleEndian;
    }

    public int ReadByte()
    {
        var bytesRead = _stream.Read(_buf, 0, 1);
        if (bytesRead == 1)
        {
            _count += 1;
            return 0xff & _buf[0];
        }

        throw new EndOfStreamException();
    }

    public int ReadInt16()
    {
        var bytesRead = _stream.Read(_buf, 0, 2);
        if (bytesRead == 2)
        {
            short value = 0;
            if (!_isLittleEndian)
            {
                value |= (short)(0xff & _buf[1]);
                value |= (short)((0xff & _buf[0]) << 8);
            }
            else
            {
                value |= (short)(0xff & _buf[0]);
                value |= (short)((0xff & _buf[1]) << 8);
            }

            _count += 2;
            return value;
        }

        throw new EndOfStreamException();
    }

    public int ReadInt32()
    {
        var bytesRead = _stream.Read(_buf, 0, 4);
        if (bytesRead == 4)
        {
            var value = 0;
            if (!_isLittleEndian)
            {
                value |= 0xff & _buf[3];
                value |= (0xff & _buf[2]) << 8;
                value |= (0xff & _buf[1]) << 16;
                value |= (0xff & _buf[0]) << 24;
            }
            else
            {
                value |= 0xff & _buf[0];
                value |= (0xff & _buf[1]) << 8;
                value |= (0xff & _buf[2]) << 16;
                value |= (0xff & _buf[3]) << 24;
            }

            _count += 4;
            return value;
        }

        throw new EndOfStreamException();
    }

    public int ReadUint16()
    {
        var bytesRead = _stream.Read(_buf, 0, 2);
        if (bytesRead == 2)
        {
            var value = 0;
            if (!_isLittleEndian)
            {
                value |= 0xff & _buf[1];
                value |= (0xff & _buf[0]) << 8;
            }
            else
            {
                value |= 0xff & _buf[0];
                value |= (0xff & _buf[1]) << 8;
            }

            _count += 2;
            return value;
        }

        throw new EndOfStreamException();
    }

    public long ReadUint32()
    {
        var bytesRead = _stream.Read(_buf, 0, 4);
        if (bytesRead == 4)
        {
            long value = 0;
            if (!_isLittleEndian)
            {
                value |= (uint)(0xff & _buf[3]);
                value |= (uint)((0xff & _buf[2]) << 8);
                value |= (uint)((0xff & _buf[1]) << 16);
                value |= (uint)((0xff & _buf[0]) << 24);
            }
            else
            {
                value |= (uint)(0xff & _buf[0]);
                value |= (uint)((0xff & _buf[1]) << 8);
                value |= (uint)((0xff & _buf[2]) << 16);
                value |= (uint)((0xff & _buf[3]) << 24);
            }

            _count += 4;
            return value;
        }

        throw new EndOfStreamException();
    }

    public byte[] ReadBytes(int n)
    {
        var array = new byte[n];
        var offset = 0;
        while (offset < n)
        {
            var r = _stream.Read(array, offset, n - offset);
            if (r == 0)
            {
                throw new EndOfStreamException();
            }

            offset += r;
        }

        _count += n;
        return array;
    }

    public int Count
    {
        get => _count;
        set => _count = value;
    }

    public void Dispose()
    {
        _stream.Dispose();
    }
}