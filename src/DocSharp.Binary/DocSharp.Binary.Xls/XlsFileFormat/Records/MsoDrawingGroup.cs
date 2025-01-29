using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.MsoDrawingGroup)] 
    public class MsoDrawingGroup : BiffRecord
    {
        public const RecordType ID = RecordType.MsoDrawingGroup;

        public MsoDrawingGroup(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // TODO: place code here
            // Record.ReadRecord(reader, 0);
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
