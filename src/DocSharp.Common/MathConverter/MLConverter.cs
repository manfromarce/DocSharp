using System.Xml;

namespace DocSharp.MathConverter;

public static class MLConverter
{
    public static string ToLaTex(XmlNode oMath) => new MLMathNode(oMath).Text;
    public static string ToLaTex(string oMathXml)
    {
        XmlDocument doc = new XmlDocument();
        doc.LoadXml(oMathXml);
        return doc.DocumentElement is XmlNode node ? new MLMathNode(node).Text : string.Empty;
    }
}
