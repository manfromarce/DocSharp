using System.IO;
using DocSharp.Binary.OfficeDrawing;

namespace DocSharp.Binary.PptFileFormat
{
    [OfficeRecord(1053)]
    public class RoundTripCompositeMasterId12Atom : Record
    {
        public uint compositeMasterId;

        public RoundTripCompositeMasterId12Atom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance)
        {
            this.compositeMasterId = this.Reader.ReadUInt32();
        }
    }

}
