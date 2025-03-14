using System.Collections.Generic;
using System.Xml;
using DocSharp.Xml;

namespace DocSharp.MathConverter;

internal class MLPropertiesNode : MLNodeBase
{
    private static readonly HashSet<string> val_tags = ["chr", "pos", "begChr", "endChr", "type"];

    private readonly string text;
    private readonly Dictionary<string, TeXNode?> inner_dict = new Dictionary<string, TeXNode?>();

    public MLPropertiesNode(XmlNode elm)
    {
        text = ProcessChildren(elm);
    }

    public override string ToString() => text;

    public string? GetAttributeValue(string name)
    {
        if (!inner_dict.ContainsKey(name))
            return null;

        return inner_dict[name]?.ToString();
    }

    protected override TeXNode? ProcessTag(string tag, XmlNode elm)
    {
        switch (tag)
        {
            case "brk": return DoBrk();
            case "chr": return DoCommon(elm);
            case "pos": return DoCommon(elm);
            case "begChr": return DoCommon(elm);
            case "endChr": return DoCommon(elm);
            case "type": return DoCommon(elm);
            default: return null;
        }
    }

    private TeXNode? DoBrk()
    {
        inner_dict["brk"] = new TeXNode(BRK);
        return inner_dict["brk"];
    }

    private TeXNode? DoCommon(XmlNode elm)
    {
        if (val_tags.Contains(elm.LocalName))
        {
            var val = elm.GetAttributeValue("m:val");
            inner_dict[elm.LocalName] = (val != null ? new TeXNode(val) : null);
        }

        return null;
    }

}
