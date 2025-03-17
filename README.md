# DocSharp

DocSharp is a pure C# library to convert between document formats without Office interop or native dependencies.

The following packages are currently available:

- DocSharp.Binary: convert Office 97-2003 binary documents (doc, xls, ppt) to OpenXML documents (docx, xlsx, pptx). This is a fork of the abandoned [b2xtranslator project](https://github.com/EvolutionJobs/b2xtranslator) which provides critical fixes. 
- DocSharp.Docx: convert DOCX to Markdown, RTF and plain text (.txt). Possible applications include generating Open XML documents in C# and exporting for other editors, or loading Word documents in a RichTextBox / RichEditBox control.
- DocSharp.Markdown: convert Markdown to DOCX using a custom Markdig renderer.

Packages can be installed via NuGet:  
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Doc?style=flat-square&label=DocSharp.Binary.Doc)](https://www.nuget.org/packages/DocSharp.Binary.Doc/) 
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Xls?style=flat-square&label=DocSharp.Binary.Xls)](https://www.nuget.org/packages/DocSharp.Binary.Xls/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Ppt?style=flat-square&label=DocSharp.Binary.Ppt)](https://www.nuget.org/packages/DocSharp.Binary.Ppt/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Docx?style=flat-square&label=DocSharp.Docx)](https://www.nuget.org/packages/DocSharp.Docx/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Markdown?style=flat-square&label=DocSharp.Markdown)](https://www.nuget.org/packages/DocSharp.Markdown/)

There is no common DOM to manipulate or generate documents, this library is mainly for conversion. However, the Docx package provides some helper methods on top of the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) that may be extended in the future.  
If your main purpose is creating documents from scratch you can consider the following libraries: [OfficeIMO](https://github.com/EvotecIT/OfficeIMO), [OpenXML-Office](https://github.com/DraviaVemal/OpenXML-Office), [ClosedXML](https://github.com/ClosedXML/ClosedXML), [ShapeCrawler](https://github.com/ShapeCrawler/ShapeCrawler), [QuestPDF](https://github.com/QuestPDF/QuestPDF), [MigraDoc](https://github.com/empira/PDFsharp).

### Supported features

- Binary formats: most doc/xls/ppt features were supported by the original project, but exceptions occurred when using .NET (rather than .NET Framework) or loading specific documents. Many errors have been fixed, but more work is needed to make the library reliable; if you find other bugs, you are welcome to open an issue (please attach a sample file if the issue only occurs for specific documents).
- DOCX, RTF, Markdown: supported elements vary depending on input and output formats, see [Supported features](https://github.com/manfromarce/DocSharp/documentation/Supported_features.MD) for an overview.

### Requirements

.NET 6, 8, 9 and .NET Framework 4.6.2 and higher are supported.

### Usage

You can refer to the project [Wiki](https://github.com/manfromarce/DocSharp/wiki) or [sample app](https://github.com/manfromarce/DocSharp/tree/main/samples/WpfApp1).

### Roadmap

- Support more elements and attributes, and fix issues on edge cases
- Reverse RTF to DOCX conversion
- Async functions/progress callback (some tasks such as downloading images referenced in Markdown may take some time)
- Improve support for right-to-left and complex script languages

### Credits

Dependencies: 
- [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)
- [Markdig](https://github.com/xoofx/markdig) - for DocSharp.Markdown only

Forked: 
- [b2xtranslator](https://github.com/EvolutionJobs/b2xtranslator)
- [markdig.docx](https://github.com/morincer/markdig.docx)

Others:
- [Html2OpenXml](https://github.com/onizet/html2openxml) for images header decoding and unit conversions.
- [dwml_cs](https://github.com/m-x-d/dwml_cs) for Office MathML conversion

### License

DocSharp is licensed under MIT license and can be used for both open source and commercial projects.  
If you find the library useful, adding a star is highly appreciated.
