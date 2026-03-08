namespace DocSharp.Renderer;

internal class QuestPdfFootnoteReference : QuestPdfInlineElement
{
    public long Id { get; set; }

    public QuestPdfFootnoteReference(long id)
    {
        this.Id = id;
    }
}
