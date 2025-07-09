using DocSharp.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class FootnoteEndnoteHelpers
{
    public static string GetFootnoteIdString(this FootnoteReferenceMark footnoteReferenceMark)
    {
        // TODO: get numbering format from FootnoteProperties or FootnoteDocumentWideProperties
        return footnoteReferenceMark.GetFootnoteId().ToStringInvariant();
    }

    public static string GetEndnoteIdString(this EndnoteReferenceMark endnoteReferenceMark)
    {
        // TODO: get numbering format from EndnoteProperties or EndnoteDocumentWideProperties
        return ListHelpers.NumberToRomanLetter(endnoteReferenceMark.GetEndnoteId(), uppercase: false);
    }

    public static string GetFootnoteIdString(this Footnote footnote)
    {
        return footnote.GetFootnoteId().ToStringInvariant();
    }

    public static string GetEndnoteIdString(this Endnote endnote)
    {
        return endnote.GetEndnoteId().ToStringInvariant();
    }

    public static string GetFootnoteIdString(this FootnoteReference footnoteReference)
    {
        return footnoteReference.GetFootnoteId().ToStringInvariant();
    }

    public static string GetEndnoteIdString(this EndnoteReference endnoteReference)
    {
        return ListHelpers.NumberToRomanLetter(endnoteReference.GetEndnoteId(), uppercase: false);
    }

    public static long GetFootnoteId(this FootnoteReferenceMark footnoteReferenceMark)
    {
        return footnoteReferenceMark.GetFirstAncestor<Footnote>()?.Id is IntegerValue id ? id.Value : 1;
    }

    public static long GetEndnoteId(this EndnoteReferenceMark endnoteReferenceMark)
    {
        return endnoteReferenceMark.GetFirstAncestor<Footnote>()?.Id is IntegerValue id ? id.Value : 1;
    }

    public static long GetFootnoteId(this Footnote footnote)
    {
        return footnote.Id is IntegerValue id ? id.Value : 1;
    }

    public static long GetEndnoteId(this Endnote endnote)
    {
        return endnote.Id is IntegerValue id ? id.Value : 1;
    }

    public static long GetFootnoteId(this FootnoteReference footnoteReference)
    {
        return footnoteReference.Id is IntegerValue id ? id.Value : 1;
    }

    public static long GetEndnoteId(this EndnoteReference endnoteReference)
    {
        return endnoteReference.Id is IntegerValue id ? id.Value : 1;
    }
}
