using DocSharp.Binary.CommonTranslatorLib;
using DocSharp.Binary.PptFileFormat;
using DocSharp.Binary.OpenXmlLib;

namespace DocSharp.Binary.PresentationMLMapping
{
    public class VbaProjectMapping : AbstractOpenXmlMapping,
        IMapping<ExOleObjStgAtom>
    {
        private VbaProjectPart _targetPart;

        public VbaProjectMapping(VbaProjectPart targetPart)
            : base(null)
        {
            this._targetPart = targetPart;
        }

        public void Apply(ExOleObjStgAtom vbaProject)
        {
            var bytes = vbaProject.DecompressData();
            this._targetPart.GetStream().Write(bytes, 0, bytes.Length);
            
        }
    }
}
