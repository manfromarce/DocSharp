using System.Collections.Generic;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class CrtMlfrtSequence : BiffRecordSequence
    {
        public List<CrtMlFrt> CrtMlFrts;

        public List<CrtMlFrtContinue> CrtMlFrtContinues;

        public CrtMlfrtSequence(IStreamReader reader)
            : base(reader)
        {
            //Spec says: CRTMLFRT = CrtMlFrt *CrtMlFrtContinue

            //Reality says: CRTMLFRT = *CrtMlFrt *CrtMlFrtContinue

            this.CrtMlFrts = new List<CrtMlFrt>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrt)
            {
                this.CrtMlFrts.Add((CrtMlFrt)BiffRecord.ReadRecord(reader));
            }

            this.CrtMlFrtContinues = new List<CrtMlFrtContinue>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.CrtMlFrtContinue)
            {
                this.CrtMlFrtContinues.Add((CrtMlFrtContinue)BiffRecord.ReadRecord(reader));
            }

        }
    }
}
