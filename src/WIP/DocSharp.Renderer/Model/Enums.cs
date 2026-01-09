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

internal enum VerticalAlignment
{
    Top,
    Center,
    Bottom
}

internal enum QuestPdfContainerType
{
    HeaderFirstPage, 
    HeaderEvenPages,
    HeaderOddOrDefault,
    FooterFirstPage, 
    FooterEvenPages,
    FooterOddOrDefault,
    Body
}