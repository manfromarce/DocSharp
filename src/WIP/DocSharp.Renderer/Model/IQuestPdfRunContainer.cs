namespace DocSharp.Renderer;

internal interface IQuestPdfRunContainer
{
    void AddSpan(QuestPdfSpan span);   
    void AddPageNumber();   
    internal IQuestPdfRunContainer CloneEmpty();
}

