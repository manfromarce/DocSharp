using System.IO;
using DocSharp.Binary.OfficeDrawing;

namespace DocSharp.Binary.PptFileFormat
{
    [OfficeRecord(4041)]
    public class Handout : RegularContainer
    {
        public Handout(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }

        public SlidePersistAtom PersistAtom;
    }

}
