using System.IO;

namespace DocSharp.Renderer.Core
{
    internal interface IImageAccessor
    {
        Stream GetImageStream(string imageId);
    }
}
