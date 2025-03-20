using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Extensions.TaskLists;

namespace Markdig.Renderers.Docx.Extensions;

public class TaskListRenderer : DocxObjectRenderer<TaskList>
{
    protected override void WriteObject(DocxDocumentRenderer renderer, TaskList obj)
    {
        WriteText(renderer, obj.Checked ? "✅" : "🔲");
        // To be improved; the list item itself is currently rendered independently.
    }
}
