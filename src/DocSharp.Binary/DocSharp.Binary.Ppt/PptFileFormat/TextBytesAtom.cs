using System.Text;
using System.IO;
using DocSharp.Binary.OfficeDrawing;

namespace DocSharp.Binary.PptFileFormat
{
    [OfficeRecord(4008)]
    public class TextBytesAtom : TextAtom
    {
        public static Encoding ENCODING = Encoding.GetEncoding("iso-8859-1");

        public TextBytesAtom(BinaryReader _reader, uint size, uint typeCode, uint version, uint instance)
            : base(_reader, size, typeCode, version, instance, ENCODING) { }
    }

}
