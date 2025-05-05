using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Extensions.TaskLists;

namespace Markdig.Renderers.Rtf.Extensions;

public class TaskListRenderer : RtfObjectRenderer<TaskList>
{
    protected override void WriteObject(RtfRenderer renderer, TaskList obj)
    {
        WriteText(renderer, obj.Checked ? "✅" : "🔲");
    }
}
