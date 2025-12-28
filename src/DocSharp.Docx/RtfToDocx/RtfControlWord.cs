namespace DocSharp.Rtf;

internal class RtfControlWord(string name) : RtfToken
{
    public string Name { get; set; } = name;
}

