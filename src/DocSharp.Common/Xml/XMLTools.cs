using System.Xml;

namespace DocSharp.Xml;

internal static class XMLTools
{
	public static XmlNode? GetChildByName(this XmlNode n, string name)
	{
		foreach (XmlNode cn in n.ChildNodes)
			if (cn.Name == name)
				return cn;

		return null;
	}

	public static string? GetAttributeValue(this XmlNode n, string name)
	{
		if (n.Attributes == null)
			return null;

		foreach (XmlAttribute attr in n.Attributes)
			if (attr.Name == name)
				return attr.Value;

		return null;
	}
}
