using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

internal class ParagraphState
{
    public ParagraphProperties? ParagraphProperties { get; set; } = null;
    public int TableNestingLevel { get; set; } = 0;

    public void Reset()
    {
        ParagraphProperties = null;
        TableNestingLevel = 0;
    }

    public ParagraphState Clone()
    {
        return new ParagraphState()
        {
            ParagraphProperties = this.ParagraphProperties == null ? null : (ParagraphProperties)this.ParagraphProperties.CloneNode(true),
            TableNestingLevel = this.TableNestingLevel
        };
    }
}
