namespace DocSharp.Rtf;

internal class RtfText(string text) : RtfToken
{
    public string Text { get; set; } = text;

    public RtfText() : this(string.Empty) { }
}

