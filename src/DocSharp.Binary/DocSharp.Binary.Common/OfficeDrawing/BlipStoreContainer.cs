using System.IO;

namespace DocSharp.Binary.OfficeDrawing
{
    [OfficeRecord(0xF001)]
    public class BlipStoreContainer : RegularContainer
    {
        public BlipStoreContainer(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance) 
        { 
            
        }
    }
}
