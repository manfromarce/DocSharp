using System.IO;

namespace DocSharp.Binary.OfficeDrawing
{
    [OfficeRecord(0xF002)]
    public class DrawingContainer : RegularContainer
    {
        public DrawingContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) 
        {
        }
    }
}
