using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DocSharp.Writers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace DocSharp.Docx;

public partial class DocxToHtmlConverter : DocxToXmlWriterBase<HtmlTextWriter>
{
    internal override void ProcessMathElement(OpenXmlElement element, HtmlTextWriter writer)
    {
        using (var stream = LoadXslTransform())
        {
            var xml = element.OuterXml;
            if (!string.IsNullOrEmpty(xml))
            {
                // Transform the OpenXML Math element to MathML using XSLT.
                using (var reader = XmlReader.Create(stream))
                {
                    var settings = new XmlReaderSettings
                    {
                        IgnoreWhitespace = true,
                        IgnoreComments = true,
                    };

                    using (var xmlReader = XmlReader.Create(new StringReader(xml), settings))
                    {
                        var doc = new XmlDocument();
                        doc.Load(xmlReader);

                        // Load the XSLT transformation.
                        var xslt = new System.Xml.Xsl.XslCompiledTransform();
                        xslt.Load(reader);

                        xslt.Transform(doc, null, writer);
                    }
                }
            }
        }
    }

    internal static Stream LoadXslTransform()
    {
        string xsl = "DocSharp.Docx.Resources.OMML2MML.XSL";

        var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(xsl);
        if (stream == null)
        {
            stream = Assembly.GetCallingAssembly().GetManifestResourceStream(xsl);
        }
        if (stream == null)
        {
            throw new FileNotFoundException($"Failed to load XSL transform from resources.");
        }
        return stream;
    }
}
