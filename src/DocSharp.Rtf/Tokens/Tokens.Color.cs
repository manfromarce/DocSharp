namespace DocSharp.Rtf.Tokens;

public class ColorTable : ControlTag
{
    public override string Name => "colortbl";
    public override TokenType Type => TokenType.HeaderTag;
}

public class Tint : ControlWord<byte>
{
    public override string Name => "tint";
    public override TokenType Type => TokenType.HeaderTag;

    public Tint(byte value) : base(value) { }
}
public class Shade : ControlWord<byte>
{
    public override string Name => "shade";
    public override TokenType Type => TokenType.HeaderTag;

    public Shade(byte value) : base(value) { }
}
public class AccentSix : ControlTag
{
    public override string Name => "accentsix";
    public override TokenType Type => TokenType.HeaderTag;
}

public class Red : ControlWord<byte>
{
    public override string Name => "red";
    public override TokenType Type => TokenType.HeaderTag;

    public Red(byte value) : base(value) { }
}

public class Green : ControlWord<byte>
{
    public override string Name => "green";
    public override TokenType Type => TokenType.HeaderTag;

    public Green(byte value) : base(value) { }
}

public class Blue : ControlWord<byte>
{
    public override string Name => "blue";
    public override TokenType Type => TokenType.HeaderTag;

    public Blue(byte value) : base(value) { }
}
