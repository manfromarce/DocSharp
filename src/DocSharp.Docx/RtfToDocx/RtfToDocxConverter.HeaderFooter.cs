using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using DocSharp.Writers;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using DocSharp.Helpers;
using System.Xml;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocSharp.Rtf;

namespace DocSharp.Docx;

public partial class RtfToDocxConverter : ITextToDocxConverter
{
    private void ProcessHeader(RtfGroup group, HeaderFooterValues type)
    {
        currentSectPr ??= CreateSectionProperties();

        // Get header reference of the specified type, if present
        var headerRefs = currentSectPr.OfType<HeaderReference>().Where(fr => fr.Type != null && fr.Type == type);
        var headerRef = headerRefs.FirstOrDefault();

        // If for some reason there are more headers of the same type, remove them
        if (headerRef != null && headerRefs.Count() > 1)
            headerRefs.Skip(1).ToList().ForEach(x => x.Remove());

        // If no header reference of the specified type was found, create it        
        headerRef ??= currentSectPr.AppendChild(new HeaderReference() { Type = type });

        // If the header part linked to this HeaderReference already exists, retrieve it and clear its contents, 
        // otherwise create a new header part.
        HeaderPart headerPart;
        if (!string.IsNullOrWhiteSpace(headerRef.Id?.Value) && 
            mainPart.TryGetPartById(headerRef.Id!.Value!, out OpenXmlPart? part) && part is HeaderPart hp)
            headerPart = hp;
        else
        {
            headerPart = mainPart.AddNewPart<HeaderPart>();
            headerRef.Id = mainPart.GetIdOfPart(headerPart);
        }
        headerPart.Header ??= new Header();
        headerPart.Header.RemoveAllChildren();
        headerPart.Header.ClearAllAttributes();

        // Add content to the header
        var oldContainer = container;
        container = headerPart.Header;
        ConvertGroup(group);
        container = oldContainer;
    }

    private void ProcessFooter(RtfGroup group, HeaderFooterValues type)
    {
        currentSectPr ??= CreateSectionProperties();

        // Get footer reference of the specified type, if present
        var footerRefs = currentSectPr.OfType<FooterReference>().Where(fr => fr.Type != null && fr.Type == type);
        var footerRef = footerRefs.FirstOrDefault();

        // If for some reason there are more footers of the same type, remove them
        if (footerRef != null && footerRefs.Count() > 1)
            footerRefs.Skip(1).ToList().ForEach(x => x.Remove());

        // If no footer reference of the specified type was found, create it        
        footerRef ??= currentSectPr.AppendChild(new FooterReference() { Type = type });

        // If the footer part linked to this FooterReference already exists, clear its content, 
        // otherwise create a new footer part.
        FooterPart footerPart;
        if (!string.IsNullOrWhiteSpace(footerRef.Id?.Value) && 
            mainPart.TryGetPartById(footerRef.Id!.Value!, out OpenXmlPart? part) && part is FooterPart fp)
            footerPart = fp;
        else
        {
            footerPart = mainPart.AddNewPart<FooterPart>();
            footerRef.Id = mainPart.GetIdOfPart(footerPart);
        }
        footerPart.Footer ??= new Footer();
        footerPart.Footer.RemoveAllChildren();
        footerPart.Footer.ClearAllAttributes();

        // Add content to the footer
        var oldContainer = container;
        container = footerPart.Footer;
        ConvertGroup(group);
        container = oldContainer;
    }
}