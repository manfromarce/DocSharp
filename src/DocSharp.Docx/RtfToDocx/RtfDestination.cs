namespace DocSharp.Rtf;

internal class RtfDestination : RtfGroup
{
	public string Name { get; }

	public RtfDestination(string name)
	{
		Name = name ?? string.Empty;
	}
}

