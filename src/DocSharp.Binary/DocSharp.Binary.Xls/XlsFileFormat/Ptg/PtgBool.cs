using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgBool : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgBool;

        public PtgBool(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            byte value = this.Reader.ReadByte();
            if ((int)value == 0)
            {
                this.Data = "FALSE";
            }
            else
            {
                this.Data = "TRUE"; 
            }
            this.Length = 2;
            this.type = PtgType.Operator;
            this.popSize = 1;
        }
    }
}
