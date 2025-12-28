namespace DocSharp.Rtf;

internal class RtfControlWordWithValue<T>(string name, T value) : RtfToken
{
    public string Name { get; set; } = name;

    public T Value { get; set; } = value;

}

