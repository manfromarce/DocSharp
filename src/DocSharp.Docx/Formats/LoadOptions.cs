using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocSharp.Docx;

public interface ILoadOptions
{
    LoadFormat Format { get; }
}

public class DocxLoadOptions : ILoadOptions
{
    public LoadFormat Format => LoadFormat.Docx;
}

public class RtfLoadOptions : ILoadOptions
{
    public LoadFormat Format => LoadFormat.Rtf;
}