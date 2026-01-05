namespace DocSharp.Renderer;

internal interface IQuestPdfRunContainer
{
    void AddSpan(QuestPdfSpan span);   
    internal IQuestPdfRunContainer CloneEmpty();
}

