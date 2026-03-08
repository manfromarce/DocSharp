namespace DocSharp.Renderer;

internal class QuestPdfBookmark : QuestPdfInlineElement
{
    public string Name { get; set; }

    public QuestPdfBookmark(string name)
    {
        this.Name = name;
    }
}
