using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

internal class FormattingState
{
    public bool Bold { get; set; }
    public bool AssociatedBold { get; set; }
    public bool Italic { get; set; }
    public bool AssociatedItalic { get; set; }
    public bool Strike { get; set; }
    public bool DoubleStrike { get; set; }
    public bool Subscript { get; set; }
    public bool Superscript { get; set; }
    public bool SmallCaps { get; set; }
    public bool AllCaps { get; set; }
    public bool Hidden { get; set; }
    public bool WebHidden { get; set; }
    public bool Emboss { get; set; }
    public bool Imprint { get; set; }
    public bool Outline { get; set; }
    public bool Shadow { get; set; }
    public bool RightToLeft { get; set; }
    public bool SnapToGrid { get; set; }
    public bool NoProof { get; set; }
    public bool ComplexScript { get; set; }

    public UnderlineValues? Underline { get; set; }
    public EmphasisMarkValues? Emphasis { get; set; }

    public Border? CharacterBorder { get; set; }
    public Shading? CharacterShading { get; set; }
    public Languages? Languages { get; set; }

    public int? FontIndex { get; set; }
    public int? AssociatedFontIndex { get; set; }
    public int? FontColorIndex { get; set; }
    public int? HighlightColorIndex { get; set; }
    public int? UnderlineColorIndex { get; set; }

    public int? FontSize { get; set; }
    public int? AssociatedFontSize { get; set; }
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
    
    // When true the last emitted break was a text-wrapping line break (from \line or \lbr).
    // Used to avoid emitting duplicate breaks when both \line and \lbr are present.
    public bool LastWasLineBreak { get; set; }

    public FormattingState Clone()
    {            
        return new FormattingState 
        { 
            Bold = this.Bold, 
            AssociatedBold = this.AssociatedBold, 
            Italic = this.Italic, 
            AssociatedItalic = this.AssociatedItalic, 
            Strike = this.Strike, 
            DoubleStrike = this.DoubleStrike,
            Subscript = this.Subscript,
            Superscript = this.Superscript,
            SmallCaps = this.SmallCaps,
            AllCaps = this.AllCaps,
            Hidden = this.Hidden,
            WebHidden = this.WebHidden,
            Emboss = this.Emboss,
            Imprint = this.Imprint,
            Outline = this.Outline,
            Shadow = this.Shadow,
            RightToLeft = this.RightToLeft,
            SnapToGrid = this.SnapToGrid,
            NoProof = this.NoProof,
            ComplexScript = this.ComplexScript,

            Underline = this.Underline,
            Emphasis = this.Emphasis,

            CharacterBorder = this.CharacterBorder,
            CharacterShading = this.CharacterShading,
            Languages = this.Languages,

            FontIndex = this.FontIndex,
            AssociatedFontIndex = this.AssociatedFontIndex,
            FontColorIndex = this.FontColorIndex,
            HighlightColorIndex = this.HighlightColorIndex,
            UnderlineColorIndex = this.UnderlineColorIndex,
            
            FontSize = this.FontSize,
            AssociatedFontSize = this.AssociatedFontSize,
            FontScaling = this.FontScaling,
            FontSpacing = this.FontSpacing,
            FitText = this.FitText,
            Kerning = this.Kerning,
            VerticalOffset = this.VerticalOffset,

            CharacterStyleIndex = this.CharacterStyleIndex,
            
            Uc = this.Uc,
            PendingAnsiSkip = this.PendingAnsiSkip,
            LastWasLineBreak = this.LastWasLineBreak
        };
    }

    public void Clear()
    {
        Bold = false;
        AssociatedBold = false;
        Italic = false;
        AssociatedItalic = false;
        Strike = false;
        DoubleStrike = false;
        Subscript = false;
        Superscript = false;
        SmallCaps = false;
        AllCaps = false;
        Hidden = false;
        WebHidden = false;
        Emboss = false;
        Imprint = false;
        Outline = false;
        Shadow = false;
        RightToLeft = false;
        SnapToGrid = false;
        NoProof = false;
        this.ComplexScript = false;

        Underline = null;
        Emphasis = null;
        CharacterBorder = null;
        CharacterShading = null;
        FontIndex = null;
        AssociatedFontIndex = null;
        FontColorIndex = null;
        HighlightColorIndex = null;
        UnderlineColorIndex = null;
        FontSize = null;
        AssociatedFontSize = null;
        FontScaling = null;
        FontSpacing = null;
        FitText = null;
        Kerning = null;
        VerticalOffset = null;
        CharacterStyleIndex = null;
        Languages = null;

        Uc = 1;
        PendingAnsiSkip = 0;
        LastWasLineBreak = false;
    }
}
