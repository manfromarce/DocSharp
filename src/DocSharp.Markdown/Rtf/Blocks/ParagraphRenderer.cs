using System.Drawing;
using DocSharp.Helpers;
using DocSharp.Markdown;
using DocumentFormat.OpenXml.Spreadsheet;
using Markdig.Syntax;

namespace Markdig.Renderers.Rtf.Blocks;

public class ParagraphRenderer : LeafBlockParagraphRendererBase<ParagraphBlock>
{
    protected override void WriteObject(RtfRenderer renderer, ParagraphBlock obj)
    {
        bool isInQuoteBlock = obj.Parent is QuoteBlock;

        renderer.RtfWriter.Append(@"\pard\plain");

        if (renderer.isInTable)
        {
            renderer.RtfWriter.Append(@"\intbl");
        }
        if (renderer.isInTableHeader)
        {
            renderer.RtfWriter.Append(@"\b");
        }

        // Write properties common to all paragraphs.
        renderer.RtfWriter.Append(@$"\sa{renderer.Settings.ParagraphSpaceAfterInTwips}\sl{renderer.Settings.LineSpacingValue}\slmult1");

        long spacing = 100;
        if (obj.Parent is ListItemBlock listItemBlock)
        {
            // Add space between bullet and text
            int firstLineIndent = 460;
            long indent = firstLineIndent * listItemBlock.FindListItemLevel();
            spacing += (indent - firstLineIndent);
            renderer.RtfWriter.Append($@"\fi-{firstLineIndent}\li{indent}");
            isInQuoteBlock |= obj.FindAncestor<QuoteBlock>() != null;
        }
        
        if (isInQuoteBlock)
        {
            // Format paragraph as a quote
            QuoteBlockRenderer.WriteQuoteFormatting(renderer, borderSpacing: spacing);
        }
        else
        {
            // Standard formatting
            renderer.RtfWriter.Append(@$"\f0\fs{renderer.Settings.DefaultFontSizeInHalfPoints}\cf1 "); 
        }

        if (obj.Parent is ListItemBlock lib)
        {
            if (lib.Parent is ListBlock lb && lb.IsOrdered)
                renderer.RtfWriter.Append($@"\contextualspace{{\pntext\f0 {lib.Order}.\tab}}{{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{{\pntxta.}}}}");
            else
                renderer.RtfWriter.Append($@"\contextualspace{{\pntext\f0 \bullet\tab}}{{\*\pn\pnlvlblt\pnf1\pnindent0\pnstart1\pndec{{\pntxtb\bullet}}}}");
        }

        RenderContents(renderer, obj);

        if (renderer.isInTableHeader)
        {
            renderer.RtfWriter.Append(@"\b0");
        }

        // Close paragraph
        if (!(obj.IsLastChild() && (renderer.isInTable || renderer.isInEndnote)))
        {
            renderer.RtfWriter.AppendLine(@"\par");
        }
    }
}
