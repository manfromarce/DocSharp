using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocSharp.Markdown;
using Markdig.Renderers;
using Markdig.Renderers.Rtf.Blocks;
using Markdig.Renderers.Rtf.Extensions;
using Markdig.Renderers.Rtf.Inlines;
using Markdig.Syntax;

namespace Markdig.Renderers.Rtf;

public class RtfRenderer : RendererBase
{
    public string? ImagesBaseUri { get; set; } = null;

    public bool SkipImages { get; set; } = false;

    public MarkdownToRtfSettings Settings;

    public StringBuilder RtfBuilder;

    internal List<string> Bookmarks = [];

    public RtfRenderer(StringBuilder rtfBuilder, MarkdownToRtfSettings settings)
    {
        Settings = settings;        
        RtfBuilder = rtfBuilder;

        // Default block renderers
        ObjectRenderers.Add(new CodeBlockRenderer());
        ObjectRenderers.Add(new ListRenderer());
        ObjectRenderers.Add(new HeadingRenderer());
        // ObjectRenderers.Add(new HtmlBlockRenderer());
        ObjectRenderers.Add(new ParagraphRenderer());
        ObjectRenderers.Add(new ListItemRenderer());
        ObjectRenderers.Add(new QuoteBlockRenderer());
        ObjectRenderers.Add(new ThematicBreakRenderer());

        // Default inline renderers
        ObjectRenderers.Add(new AutolinkInlineRenderer());
        ObjectRenderers.Add(new CodeInlineRenderer());
        ObjectRenderers.Add(new DelimiterInlineRenderer());
        ObjectRenderers.Add(new EmphasisInlineRenderer());
        ObjectRenderers.Add(new LineBreakInlineRenderer());
        ObjectRenderers.Add(new HtmlEntityInlineRenderer());
        ObjectRenderers.Add(new LinkInlineRenderer());
        ObjectRenderers.Add(new LiteralInlineRenderer());
        ObjectRenderers.Add(new HtmlInlineRenderer());

        // Extensions renderers
        ObjectRenderers.Add(new TableRenderer());
        ObjectRenderers.Add(new TaskListRenderer());
        //ObjectRenderers.Add(new MathInlineRenderer());
        //ObjectRenderers.Add(new MathBlockRenderer());
        //ObjectRenderers.Add(new FooterBlockRenderer());
        //ObjectRenderers.Add(new FigureRenderer());
        //ObjectRenderers.Add(new DefinitionListRenderer());
    }

    public override object Render(MarkdownObject markdownObject)
    {
        RtfBuilder.AppendLine(@"{\rtf1\ansi\deff0");
        RtfBuilder.Append($"{{\\fonttbl{{\\f0 {Settings.DefaultFont};}}");
        RtfBuilder.Append($"{{\\f1 {Settings.HeadingFonts[1]};}}");
        RtfBuilder.Append($"{{\\f2 {Settings.HeadingFonts[2]};}}");
        RtfBuilder.Append($"{{\\f3 {Settings.HeadingFonts[3]};}}");
        RtfBuilder.Append($"{{\\f4 {Settings.HeadingFonts[4]};}}");
        RtfBuilder.Append($"{{\\f5 {Settings.HeadingFonts[5]};}}");
        RtfBuilder.Append($"{{\\f6 {Settings.HeadingFonts[6]};}}");
        RtfBuilder.Append($"{{\\f7 {Settings.CodeFont};}}");
        RtfBuilder.AppendLineCrLf($"{{\\f8 {Settings.QuoteFont};}}}}");
        RtfBuilder.Append($"{{\\colortbl ;{Settings.DefaultTextColor.ToRtfColor()}");
        RtfBuilder.Append($"{Settings.HeadingColors[1].ToRtfColor()}");
        RtfBuilder.Append($"{Settings.HeadingColors[2].ToRtfColor()}");
        RtfBuilder.Append($"{Settings.HeadingColors[3].ToRtfColor()}");
        RtfBuilder.Append($"{Settings.HeadingColors[4].ToRtfColor()}");
        RtfBuilder.Append($"{Settings.HeadingColors[5].ToRtfColor()}");
        RtfBuilder.Append($"{Settings.HeadingColors[6].ToRtfColor()}");
        RtfBuilder.Append($"{Settings.CodeFontColor.ToRtfColor()}");
        RtfBuilder.Append($"{Settings.CodeBorderColor.ToRtfColor()}");
        RtfBuilder.Append($"{Settings.CodeBackgroundColor.ToRtfColor()}");
        RtfBuilder.Append($"{Settings.QuoteFontColor.ToRtfColor()}");
        RtfBuilder.Append($"{Settings.QuoteBorderColor.ToRtfColor()}");
        RtfBuilder.Append($"{Settings.QuoteBackgroundColor.ToRtfColor()}");
        RtfBuilder.Append(@"\red255\green255\blue0;"); // for highlighted/marked text
        RtfBuilder.Append(@"\red0\green255\blue0;"); // for inserted text
        RtfBuilder.AppendLineCrLf($"{Settings.LinkColor.ToRtfColor()}}}");
        Write(markdownObject);
        RtfBuilder.AppendLineCrLf(@"}");
        return this;
    }
}
