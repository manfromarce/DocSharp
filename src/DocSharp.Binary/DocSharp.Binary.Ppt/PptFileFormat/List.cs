using System.IO;
using DocSharp.Binary.OfficeDrawing;

namespace DocSharp.Binary.PptFileFormat
{
    // Wrongly listed in documentation as 1016
    [OfficeRecord(2000)]
    public class List : RegularContainer
    {
        public List(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

}
