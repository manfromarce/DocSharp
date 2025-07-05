using System;
using System.Diagnostics;
using DocSharp.Helpers;
using Markdig.Syntax.Inlines;

namespace Markdig.Renderers.Rtf.Inlines;

public class AutolinkInlineRenderer : RtfObjectRenderer<AutolinkInline>
{
    protected override void WriteObject(RtfRenderer renderer, AutolinkInline obj)
    {
        var uriString = obj.Url;
        var title = uriString;
        
        if (obj.IsEmail && !uriString.ToLower().StartsWith("mailto:"))
        {
            uriString = "mailto:" + uriString;
        }
        
        Uri? uri = null;

        var isAbsoluteUri = Uri.TryCreate(uriString, UriKind.Absolute, out uri);

        if (!isAbsoluteUri)
        {
            Uri.TryCreate(uriString, UriKind.Relative, out uri);
        }

        if (uri == null) return;

        renderer.RtfWriter.Append(@"{\field{\*\fldinst{HYPERLINK ");
        renderer.RtfWriter.Append(@"""" + uri + @"""}}");
        renderer.RtfWriter.Append(@"{\fldrslt{\cf16\ul ");
        renderer.RtfWriter.AppendRtfEscaped(title);
        renderer.RtfWriter.Append(@"}}}");
    }
}
