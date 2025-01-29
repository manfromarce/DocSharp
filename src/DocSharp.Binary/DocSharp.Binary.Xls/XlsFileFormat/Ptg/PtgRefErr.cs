using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgRefErr : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgRefErr;

        public PtgRefErr(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 5;
            this.Data = "#REF!";
            this.type = PtgType.Operand;
            reader.ReadBytes(4);             
        }
    }
}
