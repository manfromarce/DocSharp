using System.Diagnostics;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{

    [BiffRecord(RecordType.CrtMlFrtContinue)]
    public class CrtMlFrtContinue : BiffRecord
    {
        public const RecordType ID = RecordType.CrtMlFrtContinue;

        public FrtHeader FrtHeader;

        //An array of bytes that contains the continuation of the xmltkChain field of the CrtMlFrt record associated 
        //with this record. If the length of this record is greater than 8224 bytes, additional CrtMlFrtContinue records follow.
        public XmlTkChain XmlTkChain;

        public CrtMlFrtContinue(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.FrtHeader = new FrtHeader(reader);

            this.XmlTkChain = new XmlTkChain(reader);

            //unused
            reader.ReadBytes(4);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
