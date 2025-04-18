using DocSharp.Binary.StructuredStorage.Reader;
using System.Diagnostics;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgAdd : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgAdd;

        public PtgAdd(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 1;
            this.Data = "+";
            this.type = PtgType.Operator;
            this.popSize = 2; 
        }
    }
}
