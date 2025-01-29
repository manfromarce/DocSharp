using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.BottomMargin)]
    public class BottomMargin : BiffRecord
    {
        public const RecordType ID = RecordType.BottomMargin;
        public double value; 

        public BottomMargin(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.value = reader.ReadDouble(); 
        }
    }
}
