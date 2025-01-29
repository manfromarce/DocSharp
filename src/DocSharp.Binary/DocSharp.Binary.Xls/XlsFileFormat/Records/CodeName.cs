using System.Diagnostics;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.CodeName)] 
    public class CodeName : BiffRecord
    {
        public const RecordType ID = RecordType.CodeName;

        public XLUnicodeString codeName;

        public CodeName(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.codeName = new XLUnicodeString(reader);
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
