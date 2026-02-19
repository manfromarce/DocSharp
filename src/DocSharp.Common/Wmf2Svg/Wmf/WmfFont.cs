using DocSharp.Wmf2Svg.Gdi;

namespace DocSharp.Wmf2Svg.Wmf;

public sealed class WmfFont : WmfObject, IGdiFont
{
    public int Height { get; set; }
    public int Width { get; set; }
    public int Escapement { get; set; }
    public int Orientation { get; set; }
    public int Weight { get; set; }
    public bool IsItalic { get; set; }
    public bool IsUnderlined { get; set; }
    public bool IsStrikedOut { get; set; }
    public int Charset { get; set; }
    public int OutPrecision { get; set; }
    public int ClipPrecision { get; set; }
    public int Quality { get; set; }
    public int PitchAndFamily { get; set; }
    public string FaceName { get; set; }

    public WmfFont(int id, int height, int width, int escapement, int orientation,
        int weight, bool italic, bool underline, bool strikeout, int charset,
        int outPrecision, int clipPrecision, int quality, int pitchAndFamily,
        byte[] faceName) : base(id)
    {
        Height = height;
        Width = width;
        Escapement = escapement;
        Orientation = orientation;
        Weight = weight;
        IsItalic = italic;
        IsUnderlined = underline;
        IsStrikedOut = strikeout;
        Charset = charset;
        OutPrecision = outPrecision;
        ClipPrecision = clipPrecision;
        Quality = quality;
        PitchAndFamily = pitchAndFamily;
        FaceName = Helper.ConvertString(faceName, charset);
    }
}