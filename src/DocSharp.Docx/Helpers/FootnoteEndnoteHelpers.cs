using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class FootnoteEndnoteHelpers
{
    public static long GetFootnoteId(this FootnoteReferenceMark footnoteReferenceMark)
    {
        return footnoteReferenceMark.GetFirstAncestor<Footnote>()?.Id is IntegerValue id ? id.Value : 1;
    }

    public static long GetEndnoteId(this EndnoteReferenceMark endnoteReferenceMark)
    {
        return endnoteReferenceMark.GetFirstAncestor<Footnote>()?.Id is IntegerValue id ? id.Value : 1;
    }
}
