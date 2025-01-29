using System.Collections.Generic;
using DocSharp.Binary.Spreadsheet.XlsFileFormat.Records;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat
{
    public class ObjSequence : BiffRecordSequence
    {
        public List<Continue> Continue;

        // public Obj Obj; 
 
        public ObjSequence(IStreamReader reader)
            : base(reader)
        {
            //OBJ = Obj *Continue

            // Obj
            // this.Obj = (Obj)BiffRecord.ReadRecord(reader); 

            // *Continue
            this.Continue = new List<Continue>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Continue)
            {
                this.Continue.Add((Continue)BiffRecord.ReadRecord(reader));
            }
        }
    }
}
