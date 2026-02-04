using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

internal class FormattingState
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Strike { get; set; }
    public bool DoubleStrike { get; set; }
    public bool Subscript { get; set; }
    public bool Superscript { get; set; }
    public bool SmallCaps { get; set; }
    public bool AllCaps { get; set; }
    public bool Hidden { get; set; }
    public bool Emboss { get; set; }
    public bool Imprint { get; set; }
    public bool Outline { get; set; }
    public bool Shadow { get; set; }

    public UnderlineValues? Underline { get; set; }
    public EmphasisMarkValues? Emphasis { get; set; }

    public Border? CharacterBorder { get; set; }
    public Shading? CharacterShading { get; set; }

    public int? FontIndex { get; set; }
    public int? FontColorIndex { get; set; }
    public int? HighlightColorIndex { get; set; }
    public int? UnderlineColorIndex { get; set; }

    public int? FontSize { get; set; }
    public int? FontScaling { get; set; }
    public int? FontSpacing { get; set; }
    public int? FitText { get; set; }
    public int? Kerning { get; set; }
    public int? VerticalOffset { get; set; }

    public int? CharacterStyleIndex { get; set; }
    // Number of ANSI characters to skip after a \uN control word (per RTF \ucN)
    public int Uc { get; set; } = 1;

    // Remaining ANSI characters to skip because of a previously seen \u control word
    public int PendingAnsiSkip { get; set; }

    public FormattingState Clone()
    {            
        return new FormattingState 
        { 
            Bold = this.Bold, 
            Italic = this.Italic, 
            Strike = this.Strike, 
            DoubleStrike = this.DoubleStrike,
            Subscript = this.Subscript,
            Superscript = this.Superscript,
            SmallCaps = this.SmallCaps,
            AllCaps = this.AllCaps,
            Hidden = this.Hidden,
            Emboss = this.Emboss,
            Imprint = this.Imprint,
            Outline = this.Outline,
            Shadow = this.Shadow,

            Underline = this.Underline,
            Emphasis = this.Emphasis,

            CharacterBorder = this.CharacterBorder,
            CharacterShading = this.CharacterShading,

            FontIndex = this.FontIndex,
            FontColorIndex = this.FontColorIndex,
            HighlightColorIndex = this.HighlightColorIndex,
            UnderlineColorIndex = this.UnderlineColorIndex,
            
            FontSize = this.FontSize,
            FontScaling = this.FontScaling,
            FontSpacing = this.FontSpacing,
            FitText = this.FitText,
            Kerning = this.Kerning,
            VerticalOffset = this.VerticalOffset,

            CharacterStyleIndex = this.CharacterStyleIndex
            ,
            Uc = this.Uc,
            PendingAnsiSkip = this.PendingAnsiSkip
        };
    }

    public void Clear()
    {
        Bold = false;
        Italic = false;
        Strike = false;
        DoubleStrike = false;
        Subscript = false;
        Superscript = false;
        SmallCaps = false;
        AllCaps = false;
        Hidden = false;
        Emboss = false;
        Imprint = false;
        Outline = false;
        Shadow = false;
        Underline = null;
        Emphasis = null;
        CharacterBorder = null;
        CharacterShading = null;
        FontIndex = null;
        FontColorIndex = null;
        HighlightColorIndex = null;
        UnderlineColorIndex = null;
        FontSize = null;
        FontScaling = null;
        FontSpacing = null;
        FitText = null;
        Kerning = null;
        VerticalOffset = null;
        CharacterStyleIndex = null;
        Uc = 1;
        PendingAnsiSkip = 0;
    }
}
