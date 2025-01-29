using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Binary.Tools;
using System.IO;

namespace DocSharp.Binary.DocFileFormat
{
    public abstract class ByteStructure
    {
        protected VirtualStreamReader _reader;
        protected int _length;
        protected byte[] _rawBytes;
        public const int VARIABLE_LENGTH = int.MaxValue;

        public byte[] RawBytes => this._rawBytes;


        public ByteStructure(VirtualStreamReader reader, int length) 
        {
            this._reader = reader;
            this._length = length;

            //read the raw bytes
            if (this._length != VARIABLE_LENGTH)
            {
                this._rawBytes = this._reader.ReadBytes(this._length);
                this._reader.BaseStream.Seek(-1 * this._length, SeekOrigin.Current);
            }
        }

        public override string ToString() => 
            Utils.GetHashDump(this._rawBytes);
    }
}
