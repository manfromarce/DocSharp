using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace DocSharp.Renderer.Core
{
    internal class ImageAccessor : IImageAccessor
    {
        private readonly MainDocumentPart _mainDocumentPart;

        public ImageAccessor(MainDocumentPart mainDocumentPart)
        {
            _mainDocumentPart = mainDocumentPart;
        }

        public Stream GetImageStream(string imageId)
        {
            var imagePart = _mainDocumentPart.GetPartById(imageId);
            var stream = imagePart.GetStream();
            return stream;
        }
    }
}
