using System.Text;
using System.IO;
using DocSharp.Binary.OfficeDrawing;

namespace DocSharp.Binary.PptFileFormat
{
    [OfficeRecord(4000)]
    public class TextCharsAtom : TextAtom
    {
        public static Encoding ENCODING = Encoding.Unicode;

        public TextCharsAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance, ENCODING) { }
    }

}
