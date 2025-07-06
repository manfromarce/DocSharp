using System.Collections.Generic;

namespace DocSharp.Renderer.Models.Paragraphs.Builders
{
    internal static class StringOperations
    {
        public static IEnumerable<string> SplitToWordsAndWhitechars(this string text)
        {
            int start = 0, index;

            while ((index = text.IndexOfAny(new char[] { ' ', '\t'} , start)) != -1)
            {
                if (index - start > 0)
                    yield return text.Substring(start, index - start);
                yield return text.Substring(index, 1);
                start = index + 1;
            }

            if (start < text.Length)
            {
                yield return text.Substring(start);
            }
        }
    }
}
