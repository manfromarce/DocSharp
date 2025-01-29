using DocSharp.Binary.OfficeDrawing;
using System.IO;

namespace DocSharp.Binary.PptFileFormat
{
    [OfficeRecord(1039)]
    public class ColorMappingAtom : XmlStringAtom
    {
        public ColorMappingAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }
}
