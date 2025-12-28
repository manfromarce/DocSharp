using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Syntax;

namespace DocSharp.Markdown.Common;

public static class MarkdownElementExtensions
{
    public static int FindListItemLevel(this ListItemBlock block)
    {
        var ancestors = block.FindAncestors<ListBlock>();
        return ancestors?.Count() ?? 1;
    }

    public static IEnumerable<T>? FindAncestors<T>(this Block block) where T : ContainerBlock
    {
        var ancestor = block;
        while (ancestor != null)
        {
            ancestor = ancestor.Parent;
            if (ancestor is T ancestorAsT)
                yield return ancestorAsT;
        }
    }

    public static T? FindAncestor<T>(this Block block) where T : ContainerBlock
    {
        var ancestor = block;
        while (ancestor != null)
        {
            ancestor = ancestor.Parent;
            if (ancestor is T ancestorAsT)
                return ancestorAsT;
        }
        return null;
    }

    public static bool IsLastChild(this Block block)
    {
        return block.Parent != null && block.Parent.LastChild == block;
    }
}
