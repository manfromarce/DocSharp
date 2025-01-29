using System.Text;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Structures;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.RichTextStream)]
    public class RichTextStream : BiffRecord
    {
        public FrtHeader frtHeader;

        public uint dwCheckSum;

        public uint cb;

        public string rgb;

        public RichTextStream(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            this.frtHeader = new FrtHeader(reader);
            this.dwCheckSum = reader.ReadUInt32();
            this.cb = reader.ReadUInt32();
            var codepage = Encoding.GetEncoding("ISO-8859-1"); // windows-1252 not supported by platform
            this.rgb = codepage.GetString(reader.ReadBytes((int)this.cb));
        }
    }
}
