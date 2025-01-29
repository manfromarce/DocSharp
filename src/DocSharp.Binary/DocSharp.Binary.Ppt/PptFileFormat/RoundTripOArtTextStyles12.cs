using System.IO;
using DocSharp.Binary.OfficeDrawing;

namespace DocSharp.Binary.PptFileFormat
{
    [OfficeRecord(1059)]
    public class RoundTripOArtTextStyles12 : XmlContainer
    {
        public RoundTripOArtTextStyles12(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) { }
    }

}
