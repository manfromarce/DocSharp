# DocSharp

DocSharp is a pure C# library to convert between document formats without Office interop or native dependencies.

The following packages are currently available:

- DocSharp.Binary: convert Office 97-2003 binary documents (doc, xls, ppt) to OpenXML documents (docx, xlsx, pptx). This is a fork of the abandoned [b2xtranslator project](https://github.com/EvolutionJobs/b2xtranslator) which provides critical fixes.  
Note: pre-97 formats and XLSB are very different and not supported. For Excel files, you can also consider the [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) library to read at least plain values from more file types. 
- DocSharp.Docx: convert DOCX to RTF, HTML, Markdown and plain text (.txt). Possible applications include generating Open XML documents in C# and exporting for other editors/services, or loading Microsoft Word documents in a RichTextBox / RichEditBox control.
- DocSharp.Markdown: convert Markdown to DOCX or RTF using custom Markdig renderers.

Packages can be installed via NuGet:  
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Doc?style=flat-square&label=DocSharp.Binary.Doc)](https://www.nuget.org/packages/DocSharp.Binary.Doc/) 
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Xls?style=flat-square&label=DocSharp.Binary.Xls)](https://www.nuget.org/packages/DocSharp.Binary.Xls/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Ppt?style=flat-square&label=DocSharp.Binary.Ppt)](https://www.nuget.org/packages/DocSharp.Binary.Ppt/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Docx?style=flat-square&label=DocSharp.Docx)](https://www.nuget.org/packages/DocSharp.Docx/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Markdown?style=flat-square&label=DocSharp.Markdown)](https://www.nuget.org/packages/DocSharp.Markdown/)

The optional extra packages [DocSharp.ImageSharp](https://www.nuget.org/packages/DocSharp.ImageSharp/) and [DocSharp.SystemDrawing](https://www.nuget.org/packages/DocSharp.SystemDrawing/) allow to convert unsupported images (e.g. GIF / TIFF for DOCX -> RTF or WMF / EMF / TIFF for DOCX -> MD).  

The codebase also contains few experimental converters that are not ready and not published on NuGet yet.  

There is no common DOM to manipulate or generate documents, this library is mainly for conversion. Some helper methods on top of the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) and format-specific writers are available, but they are mostly intended for internal use; however they could be extended/improved in the future.  

You can also consider the following libraries for documents creation and manipulation: the Open XML SDK itself, [OfficeIMO](https://github.com/EvotecIT/OfficeIMO), [OpenXML-Office](https://github.com/DraviaVemal/OpenXML-Office), [ClosedXML](https://github.com/ClosedXML/ClosedXML), [ShapeCrawler](https://github.com/ShapeCrawler/ShapeCrawler), [QuestPDF](https://github.com/QuestPDF/QuestPDF), [MigraDoc](https://github.com/empira/PDFsharp).  

### Supported features

- Binary formats: most doc/xls/ppt features were supported by the original project, but exceptions occurred when using .NET (rather than .NET Framework) or loading specific documents. The most noticeable issues have been fixed, but more work is needed to make the library reliable; if you find other bugs, you are welcome to open an issue (please attach a sample file if the issue only occurs for specific documents).
- DOCX, RTF, Markdown: supported elements vary depending on input and output formats, see [Supported features](https://github.com/manfromarce/DocSharp/blob/main/documentation/Supported_features.MD) for an overview.

### Requirements

- Supported targets are .NET 8, 9, 10 and .NET Framework 4.6.2 (minimum netfx version still supported).  
- DocSharp.SystemDrawing is for Windows only (.NET Framework or net*-windows), as System.Drawing.Common is only supported on Windows; while DocSharp.ImageSharp is cross-platform for .NET 8+ (ImageSharp does not support .NET Framework).

### Usage

You can refer to the project [Wiki](https://github.com/manfromarce/DocSharp/wiki) or [sample app](https://github.com/manfromarce/DocSharp/tree/main/samples/WpfApp1).

### Roadmap

- Implement reverse RTF to DOCX conversion (⌛ started)  
- Implement DOCX renderer using QuestPDF (⌛ started)  
- Support more elements and attributes, and fix issues on edge cases
- Reduce code duplication, cleanup
- Async functions/progress callback (some tasks such as downloading images referenced in Markdown may take some time)
- Improve support for right-to-left and complex script languages
- Make converters thread-safe

### Credits

Dependencies: 
- [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)
- [Markdig](https://github.com/xoofx/markdig) - for DocSharp.Markdown
- [ImageSharp](https://github.com/SixLabors/ImageSharp) and [VectSharp](https://github.com/arklumpus/VectSharp) - for DocSharp.ImageSharp
- System.Drawing.Common and [SVG.NET](https://github.com/svg-net/SVG) - for DocSharp.SystemDrawing (supported on Windows only)

Forked: 
- [b2xtranslator](https://github.com/EvolutionJobs/b2xtranslator)
- [markdig.docx](https://github.com/morincer/markdig.docx)

Others:
- [Html2OpenXml](https://github.com/onizet/html2openxml) for images header decoding and unit conversions.  
- [dwml_cs](https://github.com/m-x-d/dwml_cs) for Office Math (OMML) to LaTex conversion  
- [addFormula2docx](https://github.com/Sun-ZhenXing/addFormula2docx) for Office Math (OMML) to MathML conversion  
- [OpenMcdf](https://github.com/ironfede/openmcdf) for better understanding Microsoft Compound format.  
- [RtfPipe](https://github.com/erdomke/RtfPipe) and [FridaysForks.RtfPipe](https://github.com/cezarypiatek/FridaysForks.RtfPipe) for part of the RTF parsing logic.  

Used in the sample app or for internal tests/comparisons (*not* dependencies when installing packages):  
- [XlsxToHtmlConverter](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter)  
- [PeachPdf](https://github.com/jhaygood86/PeachPDF)  
- [HTML-Renderer](https://github.com/ArthurHub/HTML-Renderer)
- [QuestPDF.Markdown](https://github.com/christiaanderidder/QuestPDF.Markdown)
- [ReverseMarkdown](https://github.com/mysticmind/reversemarkdown-net)
- [PDFtoImage](https://github.com/sungaila/PDFtoImage)
- [PdfToSvg.NET](https://github.com/dmester/pdftosvg.net)
- [PdfPig.Rendering.Skia](https://github.com/BobLd/PdfPig.Rendering.Skia)

### License

DocSharp is licensed under MIT license and can be used for both open source and commercial projects.  

DocSharp.ImageSharp is licensed under [Apache 2.0 license](https://www.apache.org/licenses/LICENSE-2.0.txt); ImageSharp and VectSharp have their own licenses, please visit their repositories for more information.

### Contribute

- If you know how to fix a bug, feel free to open a Pull Request.  
- To implement a new feature, please open an issue or discussion to propose it before working on the pull request.   
- If you find the library useful, adding a star is highly appreciated. Stars are a way to guide other developers towards helpful libraries and tools.
- This is a hobby project. You are welcome to donate to financially support its further development, if you wish (sponsor links for GitHub, LiberaPay, Ko-Fi, BuyMeACoffee and Thanks.dev are available in the repo page).  
