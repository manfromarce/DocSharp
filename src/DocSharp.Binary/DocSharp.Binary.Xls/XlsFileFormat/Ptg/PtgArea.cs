using System.Diagnostics;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Binary.Tools;

namespace DocSharp.Binary.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgArea : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgArea;

        public ushort rwFirst;
        public ushort rwLast;
        public ushort colFirst;
        public ushort colLast;

        public bool rwFirstRelative;
        public bool rwLastRelative;
        public bool colFirstRelative;
        public bool colLastRelative;

        public PtgArea(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 9;
            this.rwFirst = this.Reader.ReadUInt16();
            this.rwLast = this.Reader.ReadUInt16();
            this.colFirst = this.Reader.ReadUInt16();
            this.colLast = this.Reader.ReadUInt16();

            this.colFirstRelative = Utils.BitmaskToBool(this.colFirst, 0x4000);
            this.rwFirstRelative = Utils.BitmaskToBool(this.colFirst, 0x8000);
            this.colLastRelative = Utils.BitmaskToBool(this.colLast, 0x4000);
            this.rwLastRelative = Utils.BitmaskToBool(this.colLast, 0x8000);

            this.colFirst = (ushort)(this.colFirst & 0x3FFF);
            this.colLast = (ushort)(this.colLast & 0x3FFF);

            this.type = PtgType.Operand;
            this.popSize = 1;
        }
    }
}
