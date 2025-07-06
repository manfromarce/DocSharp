using System.Linq;
using Word = DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Renderer
{
    internal static class SectionXMLExtensions
    {
        public static Word.SectionMarkValues GetSectionMark(this Word.SectionProperties sectionProperties)
        {
            var st = sectionProperties.ChildsOfType<Word.SectionType>().SingleOrDefault();
            return st?.Val ?? Word.SectionMarkValues.NextPage;
        }
    }
}
