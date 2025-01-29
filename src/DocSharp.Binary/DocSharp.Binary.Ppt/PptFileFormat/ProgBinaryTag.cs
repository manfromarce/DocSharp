using DocSharp.Binary.OfficeDrawing;
using System.IO;

namespace DocSharp.Binary.PptFileFormat
{
    [OfficeRecord(5002)]
    public class ProgBinaryTag : RegularContainer
    {
        public ProgBinaryTag(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }
}
