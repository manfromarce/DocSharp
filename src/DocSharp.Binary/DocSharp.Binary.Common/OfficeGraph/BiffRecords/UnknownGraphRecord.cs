using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.OfficeGraph
{
    public class UnknownGraphRecord : OfficeGraphBiffRecord
    {
        public byte[] Content;

        public UnknownGraphRecord(IStreamReader reader, ushort id, ushort length) 
            : base(reader, (GraphRecordNumber)id, length)
        {
            this.Content = reader.ReadBytes((int)length);
        }
    }
}
