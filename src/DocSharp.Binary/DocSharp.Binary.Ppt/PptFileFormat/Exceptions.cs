using System;

namespace DocSharp.Binary.PptFileFormat
{
    public class InvalidStreamException : Exception
    {
        public InvalidStreamException()
            : base() { }

        public InvalidStreamException(string msg)
            : base(msg) { }

        public InvalidStreamException(string msg, Exception innerException)
            : base(msg, innerException) { }
    }
}
