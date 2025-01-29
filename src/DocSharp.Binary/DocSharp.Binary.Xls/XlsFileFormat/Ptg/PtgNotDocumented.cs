using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgNotDocumented : AbstractPtg
    {
        public const Ptg0x19Sub ID = Ptg0x19Sub.PtgNotDocumented;

        public PtgNotDocumented(IStreamReader reader, Ptg0x19Sub ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert((Ptg0x19Sub)this.Id == ID);
            this.Length = 4;
            this.Data = "NotDocumented";
            this.type = PtgType.Operator;
            this.popSize = 1;
            this.Reader.ReadBytes(2);
        }
    }
}



