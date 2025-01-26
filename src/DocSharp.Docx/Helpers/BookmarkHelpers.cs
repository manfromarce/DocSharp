using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class BookmarkHelpers
{
    public static string? GetBookmarkName(this BookmarkEnd bookmarkEnd)
    {
        return bookmarkEnd.GetMainDocumentPart()?.Document.Descendants<BookmarkStart>()
               .FirstOrDefault(b => b.Id == bookmarkEnd.Id)?.Name;
    }
}
