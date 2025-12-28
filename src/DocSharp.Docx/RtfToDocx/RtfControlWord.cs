namespace DocSharp.Rtf;

internal class RtfControlWord : RtfToken
{
    public string Name { get; set; }
    public int? Value { get; set; }
    public bool HasValue { get; set; }
    public bool DelimitedBySpace { get; set; }
    public RtfControlWord(string name)
    {
        Name = name;
    }
}

