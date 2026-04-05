namespace DocSharp.Rtf;

internal class RtfFontInfo
{
    public string Name { get; set; } = string.Empty;
    // \fcharsetN value found in font table entry (if any)
    public int? FCharset { get; set; }
    // \cpgN value found in font table entry (if any)
    public int? CodePage { get; set; }
}