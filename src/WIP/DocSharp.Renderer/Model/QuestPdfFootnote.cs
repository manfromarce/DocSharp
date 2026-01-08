using QuestPDF.Elements;
using QuestPDF.Fluent;
using QuestPDF.Infrastructure;

namespace DocSharp.Renderer;

internal class QuestPdfFootnote : QuestPdfContainer
{
    public int PageNumber { get; set; }
    public long Id { get; set; }

}


// public class QuestPdfFootnote2 : IDynamicComponent
// {
//     public DynamicComponentComposeResult Compose(DynamicContext context)
//     {
//         var content = context.CreateElement(element =>
//         {
//             element
//                 .Element(x => context.PageNumber % 2 == 0 ? x.AlignRight() : x.AlignLeft())
//                 .Text(text =>
//                 {
//                     text.Span("Page ");
//                     text.CurrentPageNumber();
//                 });
//         });

//         return new DynamicComponentComposeResult
//         {
//             Content = content,
//             HasMoreContent = false
//         };
//     }
// }