using System;
using System.Text;
using DocSharp.Binary.OpenXmlLib;
using DocSharp.Binary.PptFileFormat;
using System.IO;
using DocSharp.Binary.OpenXmlLib.PresentationML;
using System.Xml;

namespace DocSharp.Binary.PresentationMLMapping
{
    public class Converter
    {
        public static PresentationDocumentType DetectOutputType(PowerpointDocument ppt)
        {
            var returnType = PresentationDocumentType.Presentation;

            //ToDo: Find better way to detect macro type
            if (ppt.VbaProject != null)
            {
                returnType = PresentationDocumentType.MacroEnabledPresentation;
            }

            return returnType;
        }

        public static string GetConformFilename(string choosenFilename, PresentationDocumentType outType)
        {
            string outExt = ".pptx";
            switch (outType)
            {
                case PresentationDocumentType.MacroEnabledPresentation:
                    outExt = ".pptm";
                    break;
                case PresentationDocumentType.MacroEnabledTemplate:
                    outExt = ".potm";
                    break;
                case PresentationDocumentType.Template:
                    outExt = ".potx";
                    break;
                case PresentationDocumentType.Slideshow:
                    outExt = ".ppsx";
                    break;
                case PresentationDocumentType.MacroEnabledSlideshow:
                    outExt = ".ppsm";
                    break;
                default:
                    outExt = ".pptx";
                    break;
            }

            string inExt = Path.GetExtension(choosenFilename);
            if (inExt != null)
            {
                return choosenFilename.Replace(inExt, outExt);
            }
            else
            {
                return choosenFilename + outExt;
            }
        }

        public static void Convert(PowerpointDocument ppt, PresentationDocument pptx)
        {
            // Setup the writer
            var xws = new XmlWriterSettings();
            xws.OmitXmlDeclaration = false;
            xws.CloseOutput = true;
            xws.Encoding = Encoding.UTF8;
            xws.ConformanceLevel = ConformanceLevel.Document;

            // Setup the context
            var context = new ConversionContext(ppt)
            {
                WriterSettings = xws,
                Pptx = pptx
            };

            // Write presentation.xml
            ppt.Convert(new PresentationPartMapping(context));

            // TODO: write core/app properties (author, title, ...)

            //AppMapping app = new AppMapping(pptx.AddAppPropertiesPart(), xws);
            //app.Apply(null);

            //CoreMapping core = new CoreMapping(pptx.AddCoreFilePropertiesPart(), xws);
            //core.Apply(null);
        }
    }
}
