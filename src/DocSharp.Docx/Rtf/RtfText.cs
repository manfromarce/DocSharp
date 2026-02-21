namespace DocSharp.Rtf;

internal class RtfText : RtfToken
{
    public string Text { get; set; }

    public RtfText(string text)
    {
        Text = text ?? string.Empty;
    }

    public RtfText() : this(string.Empty) { }
}
