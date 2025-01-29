using DocSharp.Binary.OfficeDrawing;
using System.IO;

namespace DocSharp.Binary.PptFileFormat
{
    [OfficeRecord(5000)]
    public class ProgTags : RegularContainer
    {
        public ProgTags(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) {       
        }
    }
}
