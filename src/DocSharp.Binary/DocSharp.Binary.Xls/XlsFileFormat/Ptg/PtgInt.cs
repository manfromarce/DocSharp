using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgInt: AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgInt;

        public PtgInt(IStreamReader reader, PtgNumber ptgid) : 
            base(reader,ptgid) 
        {
            Debug.Assert(this.Id == ID);
            this.Length = 3; 
            this.Data = reader.ReadUInt16().ToString();

            this.type = PtgType.Operand;
            this.popSize = 1; 
        }
    }   
}
