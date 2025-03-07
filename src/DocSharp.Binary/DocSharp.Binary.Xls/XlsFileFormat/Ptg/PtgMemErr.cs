using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgMemErr : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgMemErr;

        public PtgMemErr(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 7;

            this.Reader.ReadBytes(6); 

            this.type = PtgType.Operand;
            this.popSize = 1;
        }
    }
}
