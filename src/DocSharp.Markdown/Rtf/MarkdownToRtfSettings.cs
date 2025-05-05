using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocSharp.Markdown;

public class MarkdownToRtfSettings
{
    public string DefaultFont = "Arial";
    public int DefaultFontSize = 12;
    public Color DefaultTextColor = Color.Black;

    public Dictionary<int, string> HeadingFonts = new()
    {
        [1] = "Arial",
        [2] = "Arial",
        [3] = "Arial",
        [4] = "Arial",
        [5] = "Arial",
        [6] = "Arial",
    };
    public Dictionary<int, int> HeadingFontSizes = new()
    {
        [1] = 36,
        [2] = 24,
        [3] = 20,
        [4] = 16,
        [5] = 14,
        [6] = 12,
    };
    public Dictionary<int, bool> HeadingIsBold = new()
    {
        [1] = true,
        [2] = true,
        [3] = true,
        [4] = true,
        [5] = true,
        [6] = true,
    };
    public Dictionary<int, Color> HeadingColors = new()
    {
        // Default Word heading color
        //[1] = Color.FromArgb(15, 71, 97),
        //[2] = Color.FromArgb(15, 71, 97),
        //[3] = Color.FromArgb(15, 71, 97),
        //[4] = Color.FromArgb(15, 71, 97),
        //[5] = Color.FromArgb(15, 71, 97),
        //[6] = Color.FromArgb(15, 71, 97),

        // Black (consistent with other Markdown renderers)
        [1] = Color.Black,
        [2] = Color.Black,
        [3] = Color.Black,
        [4] = Color.Black,
        [5] = Color.Black,
        [6] = Color.Black,
    };

    public string QuoteFont = "Arial";
    public Color QuoteFontColor = Color.FromArgb(118, 113, 113);
    public int QuoteFontSize = 12;
    public int QuoteIndent = 10;
    public decimal QuoteBorderWidth = 2.25m;
    public Color QuoteBorderColor = Color.FromArgb(166, 166, 166);
    public Color QuoteBackgroundColor = Color.Transparent;

    public string CodeFont = "Courier New";
    public int CodeFontSize = 12;
    public Color CodeFontColor = Color.Black;
    public Color CodeBorderColor = Color.FromArgb(128, 128, 128);
    public decimal CodeBorderWidth = 0.5m;
    public Color CodeBackgroundColor = Color.FromArgb(217, 217, 217);

    /// <summary>
    /// Paragraph spacing in points. Default is 8.
    /// </summary>
    public int ParagraphSpaceAfter = 8;
    
    /// <summary>
    /// Line spacing in lines. Default is 1.15 lines.
    /// </summary>
    public decimal LineSpacing = 1.15m;

    /// <summary>
    /// Hyperlink color. Default is blue.
    /// </summary>
    public Color LinkColor = Color.Blue;

    internal long ParagraphSpaceAfterInTwips => ParagraphSpaceAfter * 20;
    internal long LineSpacingValue => (long)Math.Round(LineSpacing * 240m, 0);
    internal long CodeBorderWidthInTwips => (long)Math.Round(CodeBorderWidth * 20m, 0);
    internal long QuoteBorderWidthInTwips => (long)Math.Round(QuoteBorderWidth * 20m, 0);
    internal int DefaultFontSizeInHalfPoints => DefaultFontSize * 2;
    internal int CodeFontSizeInHalfPoints => CodeFontSize * 2;
    internal int QuoteFontSizeInHalfPoints => QuoteFontSize * 2;

    internal int GetHeadingFontSizeInHalfPoints(int headingLevel)
    {
        return this.HeadingFontSizes[headingLevel] * 2;
    }
}
