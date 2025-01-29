using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgLe : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgLe;

        public PtgLe(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 1;
            this.Data = "<=";
            this.type = PtgType.Operator;
            this.popSize = 2;
        }
    }
}
