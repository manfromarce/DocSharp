using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// NOTE: This is STANDARDWIDTH in the previously released spec
    /// </summary>
    [BiffRecord(RecordType.DxGCol)] 
    public class DxGCol : BiffRecord
    {
        public const RecordType ID = RecordType.DxGCol;

        public DxGCol(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // TODO: place code here
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
