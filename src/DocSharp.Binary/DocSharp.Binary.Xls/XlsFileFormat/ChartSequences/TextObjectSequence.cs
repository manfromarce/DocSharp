using System.Collections.Generic;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class TextObjectSequence : BiffRecordSequence
    {
        public TxO TxO;

        public List<Continue> Continue; 

        public TextObjectSequence(IStreamReader reader)
            : base(reader)
        {
            //TEXTOBJECT = TxO *Continue
            // TxO
            this.TxO = (TxO)BiffRecord.ReadRecord(reader);

            // Continue
            this.Continue = new List<Continue>(); 
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                this.Continue.Add((Continue)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
