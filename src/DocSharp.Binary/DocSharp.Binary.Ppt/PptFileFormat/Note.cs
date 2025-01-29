using System.IO;
using DocSharp.Binary.OfficeDrawing;

namespace DocSharp.Binary.PptFileFormat
{
    [OfficeRecord(1008)]
    public class Note : RegularContainer
    {
        public Note(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }

        public SlidePersistAtom PersistAtom;
    }

}
