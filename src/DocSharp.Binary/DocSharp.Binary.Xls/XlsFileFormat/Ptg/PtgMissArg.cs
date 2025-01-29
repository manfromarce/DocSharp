using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgMissArg : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgMissArg;

        public PtgMissArg(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 1;
            this.Data = " ";
            this.type = PtgType.Operand;
            this.popSize = 1; 
        }
    }
}

