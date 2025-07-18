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
        RtfWriter.WriteRtfHeader();

        // A4 paper size in twips (1/1440 inch)
        RtfWriter.WriteLine(@"\paperw11906\paperh16780");

        // Enable endnotes
        RtfWriter.WriteLine(@"\enddoc\aenddoc");

        // Font table
        RtfWriter.Write($"{{\\fonttbl{{\\f0 {Settings.DefaultFont};}}");
        RtfWriter.Write($"{{\\f1 {Settings.HeadingFonts[1]};}}");
        RtfWriter.Write($"{{\\f2 {Settings.HeadingFonts[2]};}}");
        RtfWriter.Write($"{{\\f3 {Settings.HeadingFonts[3]};}}");
        RtfWriter.Write($"{{\\f4 {Settings.HeadingFonts[4]};}}");
        RtfWriter.Write($"{{\\f5 {Settings.HeadingFonts[5]};}}");
        RtfWriter.Write($"{{\\f6 {Settings.HeadingFonts[6]};}}");
        RtfWriter.Write($"{{\\f7 {Settings.CodeFont};}}");
        RtfWriter.WriteLine($"{{\\f8 {Settings.QuoteFont};}}}}");

        // Color table
        RtfWriter.Write("{\\colortbl ;");
        RtfWriter.Write(Settings.DefaultTextColor);
        RtfWriter.Write(Settings.HeadingColors[1]);
        RtfWriter.Write(Settings.HeadingColors[2]);
        RtfWriter.Write(Settings.HeadingColors[3]);
        RtfWriter.Write(Settings.HeadingColors[4]);
        RtfWriter.Write(Settings.HeadingColors[5]);
        RtfWriter.Write(Settings.HeadingColors[6]);
        RtfWriter.Write(Settings.CodeFontColor);
        RtfWriter.Write(Settings.CodeBorderColor);
        RtfWriter.Write(Settings.CodeBackgroundColor);
        RtfWriter.Write(Settings.QuoteFontColor);
        RtfWriter.Write(Settings.QuoteBorderColor);
        RtfWriter.Write(Settings.QuoteBackgroundColor);
        RtfWriter.Write(System.Drawing.Color.Yellow); // for highlighted/marked text
        RtfWriter.Write(System.Drawing.Color.LightGreen); // for inserted text
        RtfWriter.Write(System.Drawing.Color.FromArgb(217, 217, 217)); // for table header background
        RtfWriter.Write(Settings.LinkColor);
        RtfWriter.WriteLine("}");

        // Enable endnotes
        RtfWriter.WriteLine(@"\sectd\endnhere");

        // Add content
        Write(markdownObject);

        // End of RTF document
        RtfWriter.WriteLine(@"}");
        return this;
    }
}
