using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocSharp.Helpers;
using DocSharp.Markdown;
using DocSharp.Writers;
using Markdig.Extensions.Footnotes;
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

    public RtfStringWriter RtfWriter;

    internal List<string> Bookmarks = [];

    internal bool isInTable = false;
    internal bool isInTableHeader = false;
    internal bool isInEndnote = false;

    public RtfRenderer(RtfStringWriter rtfBuilder, MarkdownToRtfSettings settings)
    {
        Settings = settings;        
        RtfWriter = rtfBuilder;

        // Default block renderers
        ObjectRenderers.Add(new CodeBlockRenderer());
        ObjectRenderers.Add(new ListRenderer());
        ObjectRenderers.Add(new HeadingRenderer());
        // ObjectRenderers.Add(new HtmlBlockRenderer()); // Raw HTML is not supported
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
        ObjectRenderers.Add(new FooterBlockRenderer());
        ObjectRenderers.Add(new FootnoteGroupRenderer());
        ObjectRenderers.Add(new FootnoteLinkRenderer());
        // ObjectRenderers.Add(new DefinitionListRenderer());
        // ObjectRenderers.Add(new FigureRenderer());
        // ObjectRenderers.Add(new MathInlineRenderer()); // LaTex blocks are not supported
        // ObjectRenderers.Add(new MathBlockRenderer());
    }

    public override object Render(MarkdownObject markdownObject)
    {
        RtfWriter.Append(@"{\rtf1\ansi\deff0");

        // A4 paper size in twips (1/1440 inch)
        RtfWriter.AppendLine(@"\paperw11906\paperh16780");

        // Enable endnotes
        RtfWriter.AppendLine(@"\enddoc\aenddoc");

        // Font table
        RtfWriter.Append($"{{\\fonttbl{{\\f0 {Settings.DefaultFont};}}");
        RtfWriter.Append($"{{\\f1 {Settings.HeadingFonts[1]};}}");
        RtfWriter.Append($"{{\\f2 {Settings.HeadingFonts[2]};}}");
        RtfWriter.Append($"{{\\f3 {Settings.HeadingFonts[3]};}}");
        RtfWriter.Append($"{{\\f4 {Settings.HeadingFonts[4]};}}");
        RtfWriter.Append($"{{\\f5 {Settings.HeadingFonts[5]};}}");
        RtfWriter.Append($"{{\\f6 {Settings.HeadingFonts[6]};}}");
        RtfWriter.Append($"{{\\f7 {Settings.CodeFont};}}");
        RtfWriter.AppendLine($"{{\\f8 {Settings.QuoteFont};}}}}");

        // Color table
        RtfWriter.Append($"{{\\colortbl ;{Settings.DefaultTextColor.ToRtfColor()}");
        RtfWriter.Append($"{Settings.HeadingColors[1].ToRtfColor()}");
        RtfWriter.Append($"{Settings.HeadingColors[2].ToRtfColor()}");
        RtfWriter.Append($"{Settings.HeadingColors[3].ToRtfColor()}");
        RtfWriter.Append($"{Settings.HeadingColors[4].ToRtfColor()}");
        RtfWriter.Append($"{Settings.HeadingColors[5].ToRtfColor()}");
        RtfWriter.Append($"{Settings.HeadingColors[6].ToRtfColor()}");
        RtfWriter.Append($"{Settings.CodeFontColor.ToRtfColor()}");
        RtfWriter.Append($"{Settings.CodeBorderColor.ToRtfColor()}");
        RtfWriter.Append($"{Settings.CodeBackgroundColor.ToRtfColor()}");
        RtfWriter.Append($"{Settings.QuoteFontColor.ToRtfColor()}");
        RtfWriter.Append($"{Settings.QuoteBorderColor.ToRtfColor()}");
        RtfWriter.Append($"{Settings.QuoteBackgroundColor.ToRtfColor()}");
        RtfWriter.Append(@"\red255\green255\blue0;"); // for highlighted/marked text
        RtfWriter.Append(@"\red0\green255\blue0;"); // for inserted text
        RtfWriter.Append(@"\red217\green217\blue217;"); // for table header background
        RtfWriter.AppendLine($"{Settings.LinkColor.ToRtfColor()}}}");

        // Enable endnotes
        RtfWriter.AppendLine(@"\sectd\endnhere");

        // Add content
        Write(markdownObject);

        // End of RTF document
        RtfWriter.AppendLine(@"}");
        return this;
    }
}
