using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class PicfSequence : BiffRecordSequence
    {
        public Begin Begin;

        public PicF PicF;

        public End End; 

        public PicfSequence(IStreamReader reader)
            : base(reader)
        {
            // PICF = Begin PicF End

            // Begin
            this.Begin = (Begin)BiffRecord.ReadRecord(reader);

            // PicF 
            this.PicF = (PicF)BiffRecord.ReadRecord(reader);

            // End 
            this.End = (End)BiffRecord.ReadRecord(reader); 

        }
    }
}
