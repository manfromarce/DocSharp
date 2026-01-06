namespace DocSharp.Renderer;

internal enum SubSuperscript
{
    Normal,
    Subscript,
    Superscript
}

internal enum CapsType
{
    Normal,
    SmallCaps,
    AllCaps
}

internal enum ParagraphAlignment
{
    Left,
    Center,
    Right,
    Justify,
    Start,
    End
}

internal enum UnderlineStyle
{
    None,
    Solid,
    Dashed,
    Dotted,
    Double,
    Wavy
}

internal enum StrikethroughStyle
{
    None,
    Single,
    Double
}

internal enum QuestPdfContainerType
{
    HeaderFooterFirstPage, 
    HeaderFooterEvenPages,
    HeaderFooterOddOrDefault,
    Body
}