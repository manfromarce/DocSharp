using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Markdown.Common;

internal static class BookmarkHelpers
{
    internal static string GetBookmarkName(string text)
    {
        // Remove symbols and punctuation marks
        char[] normalized = text.Where(c => char.IsLetterOrDigit(c) ||
                                            c == ' ').ToArray();
                                            //char.IsWhiteSpace(c)).ToArray();

        // Trim leading/trailing spaces and replace other space with dash (-)
        return new string(normalized).Trim().Replace(" ", "-").ToLower();
    }
}
