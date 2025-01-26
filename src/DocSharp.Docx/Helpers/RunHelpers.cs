using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public static class RunHelpers
{
    public static void SetStyle(this Run run, string? styleId)
    {
        if (styleId == null) return;

        run.RunProperties ??= new RunProperties();
        run.RunProperties.RunStyle ??= new RunStyle();
        run.RunProperties.RunStyle.Val = styleId;
    }

    public static RunProperties GetOrCreateProperties(this Run run)
    {
        if (run.RunProperties == null)
        {
            run.RunProperties = new RunProperties();
        }

        return run.RunProperties;
    }
}
