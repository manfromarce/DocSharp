using System;
using System.Text;
using System.IO;

namespace DocSharp.Docx;

// Small stream adapter that writes byte chunks as text into a TextWriter.
// The Base64 data produced by ToBase64Transform is ASCII-only, so Encoding.ASCII is appropriate.
internal sealed class TextWriterStream : Stream
{
    private readonly TextWriter _writer;
    public TextWriterStream(TextWriter writer) => _writer = writer ?? throw new ArgumentNullException(nameof(writer));

    public override bool CanRead => false;
    public override bool CanSeek => false;
    public override bool CanWrite => true;
    public override long Length => throw new NotSupportedException();
    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }

    public override void Flush() => _writer.Flush();

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count)
    {
        // Base64 output is ASCII; use ASCII decoding which is slightly cheaper and unambiguous here.
        _writer.Write(Encoding.ASCII.GetString(buffer, offset, count));
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _writer.Dispose();
        }
        base.Dispose(disposing);
    }
}
