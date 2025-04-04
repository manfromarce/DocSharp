using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgPercent : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgPercent;

        public PtgPercent(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 1;
            this.type = PtgType.Operator;
            this.popSize = 1;
            this.Data = "%"; 
        }
    }
}
