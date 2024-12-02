// Copyright (c) Nicolas Musset. All rights reserved.
// This file is licensed under the MIT license.
// See the LICENSE.md file in the project root for more information.

using Markdig.Syntax;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Docx.Inlines
{
    /// <summary>
    /// A XAML renderer for a <see cref="LiteralInline"/>.
    /// </summary>
    /// <seealso cref="XamlObjectRenderer{TObject}" />
    public class LiteralInlineRenderer : DocxObjectRenderer<LiteralInline>
    {
        protected override void WriteObject(DocxDocumentRenderer renderer, LiteralInline obj)
        {
            if (obj.Content.IsEmpty)
                return;

            WriteText(renderer, obj.Content.ToString());
        }
    }
}
