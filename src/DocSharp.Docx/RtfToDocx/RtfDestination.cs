namespace DocSharp.Rtf;

internal class RtfDestination : RtfGroup
{
	public string Name { get; }

	// Set to true if the destination starts with '*'
	public bool IsIgnorable { get; }

	// Special destinations such as pnseclvl can have a numeric parameter
	public int? Value { get; set; }
    public bool HasValue { get; set; }

	public RtfDestination(string name, bool isIgnorable = false)
	{
		Name = name ?? string.Empty;
		IsIgnorable = isIgnorable;
	}
}

