namespace DocSharp.Rtf;

internal class RtfChar : RtfToken
{
    public byte CharCode { get; set; }
    public RtfChar(byte charCode)
    {
        CharCode = charCode;
    }
}

