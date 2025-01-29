using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class DatSequence : BiffRecordSequence
    {
        public Dat Dat;
        public Begin Begin;
        public LdSequence LdSequence;
        public End End;

        public DatSequence(IStreamReader reader)
            : base(reader)
        {
            // DAT = Dat Begin LD End

            this.Dat = (Dat)BiffRecord.ReadRecord(reader);
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);
            this.LdSequence = new LdSequence(reader);
            this.End = (End)BiffRecord.ReadRecord(reader);
        }
    }
}
