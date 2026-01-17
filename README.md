# DocSharp

DocSharp is a pure C# library to convert between document formats without Office interop or native dependencies (except for some special packages, see [requirements](#requirements)).

The following packages are currently available:

- DocSharp.Binary: convert Office 97-2003 binary documents (doc, xls, ppt) to OpenXML documents (docx, xlsx, pptx). This is a fork of the abandoned [b2xtranslator project](https://github.com/EvolutionJobs/b2xtranslator) which provides critical fixes.  
Note: pre-97 formats and XLSB are very different and not supported.
- DocSharp.Docx: convert DOCX to RTF, HTML, Markdown and plain text (.txt). Possible applications include generating Open XML documents in C# and exporting for other editors/services, or loading Microsoft Word documents in a RichTextBox / RichEditBox control.
- DocSharp.Markdown: convert Markdown to DOCX or RTF using custom Markdig renderers.

Packages can be installed via NuGet:  
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Doc?style=flat-square&label=DocSharp.Binary.Doc)](https://www.nuget.org/packages/DocSharp.Binary.Doc/) 
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Xls?style=flat-square&label=DocSharp.Binary.Xls)](https://www.nuget.org/packages/DocSharp.Binary.Xls/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Ppt?style=flat-square&label=DocSharp.Binary.Ppt)](https://www.nuget.org/packages/DocSharp.Binary.Ppt/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Docx?style=flat-square&label=DocSharp.Docx)](https://www.nuget.org/packages/DocSharp.Docx/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Markdown?style=flat-square&label=DocSharp.Markdown)](https://www.nuget.org/packages/DocSharp.Markdown/)

The optional extra packages [DocSharp.ImageSharp](https://www.nuget.org/packages/DocSharp.ImageSharp/), [DocSharp.SystemDrawing](https://www.nuget.org/packages/DocSharp.SystemDrawing/), DocSharp.MagickNET (not published yet) allow to convert unsupported images (e.g. GIF / TIFF for DOCX -> RTF or WMF / EMF / TIFF for DOCX -> Markdown/HTML). Each of these has pros and cons, the choice depends on your requirements. More information can be found in the [Wiki](https://github.com/manfromarce/DocSharp/wiki/Convert-images).

The codebase also contains few experimental converters that are not ready and not published on NuGet yet:  
- RTF to DOCX converter class in the DocSharp.Docx project
- DocSharp.Renderer: provides DOCX to PDF/images/SVG/XPS conversion using [QuestPDF](https://github.com/QuestPDF/QuestPDF).  
- DocSharp.Ebook: provides basic EPUB to DOCX (via HTML) conversion.

There is no common DOM to manipulate or generate documents, this library is mainly for conversion. Some helper methods on top of the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) and format-specific writers are available, but they are mostly intended for internal use; however they could be extended/improved in the future.  
You can consider using the Open XML SDK itself or other <a href="#recommended_libraries">recommended libraries</a> for documents creation and manipulation. Some of these are used in the sample app to test two-steps conversions, compare results, or generate documents in multiple formats with the same code.  
DocSharp provides methods to accept/return a WordprocessingDocument directly (in addition to file path / Stream / byte array), and a SaveTo extension method for WordprocessingDocument.

### Supported features

- Binary formats: most doc/xls/ppt features were supported by the original project, but exceptions occurred when using .NET (rather than .NET Framework) or loading specific documents. The most noticeable issues have been fixed, but more work is needed to make the library reliable; if you find other bugs, you are welcome to open an issue (please attach a sample file if the issue only occurs for specific documents).
- DOCX, RTF, Markdown: supported elements vary depending on input and output formats, see [Supported features](https://github.com/manfromarce/DocSharp/blob/main/documentation/Supported_features.MD) for an overview.

<a id="Requirements"></a>

### Requirements

- Supported targets are .NET 8, 9, 10 and .NET Framework 4.6.2 (minimum netfx version still supported).  
- DocSharp.SystemDrawing is for Windows only (.NET Framework or net*-windows), as System.Drawing.Common is based on GDI+ and only supported on Windows since .NET 6.
- DocSharp.ImageSharp is cross-platform for .NET 8+, as ImageSharp is fully managed C# code but does not support .NET Framework.
- DocSharp.MagickNET is cross-platform for both .NET and .NET Framework, but Magick.NET bundles many native libraries that might not work on non-desktop platforms (Android / iOS / WASM)
- DocSharp.Renderer depends on QuestPDF, which currently supports Windows x64 / x86, macOS x64 / ARM64, Linux x64 / ARM64. Windows ARM64, Android, iOS are not supported yet, due to a custom Skia build. Plus, the XPS generation is only supported on Windows. 

### Usage

You can refer to the project [Wiki](https://github.com/manfromarce/DocSharp/wiki) or [sample app](https://github.com/manfromarce/DocSharp/tree/main/samples/WpfApp1).

### Roadmap

- Finish and publish experimental converters
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
- System.Drawing.Common and [SVG.NET](https://github.com/svg-net/SVG) - for DocSharp.SystemDrawing
- [CoreJ2K](https://github.com/cinderblocks/CoreJ2K) - for JPEG2000 support in both DocSharp.ImageSharp and DocSharp.System.Drawing
- [Magick.NET-Q8-AnyCPU](https://github.com/dlemstra/Magick.NET) - for DocSharp.MagickNET
- [QuestPDF](https://github.com/QuestPDF/QuestPDF) - for DocSharp.Renderer
- [EpubCore](https://github.com/Pennable/EpubCore), [Html2OpenXml](https://github.com/onizet/html2openxml), [PreMailer.Net](https://github.com/milkshakesoftware/PreMailer.Net), [AngleSharp](https://github.com/AngleSharp/AngleSharp) - for DocSharp.Epub (AngleSharp is a dependency of Html2OpenXml and PreMailer.Net)

Forked: 
- [b2xtranslator](https://github.com/EvolutionJobs/b2xtranslator)
- [markdig.docx](https://github.com/morincer/markdig.docx)

Others (credits for parts of the logic):
- [Html2OpenXml](https://github.com/onizet/html2openxml) for images header decoding and unit conversions.  
- [dwml_cs](https://github.com/m-x-d/dwml_cs) for Office Math (OMML) to LaTex conversion  
- [addFormula2docx](https://github.com/Sun-ZhenXing/addFormula2docx) for Office Math (OMML) to MathML conversion  
- [RtfPipe](https://github.com/erdomke/RtfPipe), [FridaysForks.RtfPipe](https://github.com/cezarypiatek/FridaysForks.RtfPipe), [RtfConverter](https://github.com/jokecamp/RtfConverter) for part of the RTF parsing logic.  
- [ExcelNumberFormat](https://github.com/andersnm/ExcelNumberFormat) for Excel format strings parsing logic.

<a id="recommended_libraries"></a>
Other recommended libraries (some of these are used in the sample app, *not* dependencies when installing packages):  
- Read, write, manipulate docuents: 
    + [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) - DOCX, XLSX, PPTX
    + [OfficeIMO](https://github.com/EvotecIT/OfficeIMO) - DOCX, XLSX, PPTX, Markdown, CSV; can also merge, compare and convert some formats
    + [Clippit](https://github.com/sergey-tihon/Clippit) - DOCX, XLSX, PPTX; can also merge, compare and convert some formats
    + [Openize.OpenXML-SDK](https://github.com/openize-com/openize-open-xml-sdk-net) - DOCX, XLSX, PPTX
    + [ShapeCrawler](https://github.com/ShapeCrawler/ShapeCrawler) - PPTX; can also render slides to images
    + [ClosedXML](https://github.com/ClosedXML/ClosedXML) - XLSX
    + [Sylvan.Data.Excel](https://github.com/MarkPflug/Sylvan.Data.Excel) - XLSX, XLS, XLSB
    + [NPOI](https://github.com/nissl-lab/npoi) - DOCX, XLSX, XLS; partial port of Apache POI
    + [FluentNPOI](https://github.com/HouseAlwaysWin/FluentNPOI) - XLSX, XLS; HTML/PDF export
- Extract data: 
    + [GustavoHennig/b2xtranslator](https://github.com/GustavoHennig/b2xtranslator) - DOC prior to Office 97
    + [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) - XLS (pre-97 too), XLSB, XLSX, CSV
    + [PdfPig](https://github.com/UglyToad/PdfPig), [Tabula.Csv](https://github.com/BobLd/tabula-sharp) - PDF
    + [OpenMcdf](https://github.com/ironfede/openmcdf) - Microsoft Compound format
- Generate documents: 
    + PDF, XPS, SVG, images: [QuestPDF](https://github.com/QuestPDF/QuestPDF), [FossPDF.NET](https://github.com/FossPDF/FossPDF.Net)
    + PDF and RTF: [PdfSharp / MigraDoc](https://github.com/empira/PDFsharp)
    + PDF and XLSX: [PdfRpt.Core](https://github.com/VahidN/PdfReport.Core)
    + PDF, RTF, HTML: [iTextSharp.LGPLv2.Core](https://github.com/VahidN/iTextSharp.LGPLv2.Core)
    + DOCX: [SharpDocx](https://github.com/egonl/SharpDocx), [DocxTemplater](https://github.com/Amberg/DocxTemplater), [MiniWord](https://github.com/mini-software/MiniWord)
    + XLSX: [MiniExcel](https://github.com/mini-software/MiniExcel), [ClosedXML.Report](https://github.com/ClosedXML/ClosedXML.Report)
    + XLSX, ODS, CSV: [FreeDataExports](https://github.com/ryankueter/FreeDataExports)
- Convert or render documents: 
    + XLSX: [XlsxToHtmlConverter](https://github.com/Fei-Sheng-Wu/XlsxToHtmlConverter)  
    + HTML to PDF/images: [HTML-Renderer](https://github.com/ArthurHub/HTML-Renderer), [PeachPdf](https://github.com/jhaygood86/PeachPDF), [Puppeteer Sharp](https://github.com/hardkoded/puppeteer-sharp), [Westwind.WebView](https://github.com/RickStrahl/Westwind.WebView)
    + HTML to DOCX: [Html2OpenXml](https://github.com/onizet/html2openxml)
    + HTML to Markdown: [ReverseMarkdown](https://github.com/mysticmind/reversemarkdown-net)
    + Markdown to HTML: [Markdig](https://github.com/xoofx/markdig)
    + AsciiDoc to HTML: [NAsciidoc](https://github.com/rmannibucau/NAsciidoc)
    + Markdown to PDF: [QuestPDF.Markdown](https://github.com/christiaanderidder/QuestPDF.Markdown), [MarkdownToPdf](https://github.com/geertjanthomas/MarkdownToPdf), [VectSharp.Markdown + VectSharp.PDF](https://github.com/arklumpus/VectSharp)
    + PDF to images/SVG: [PDFtoImage](https://github.com/sungaila/PDFtoImage), [PdfPig.Rendering.Skia](https://github.com/BobLd/PdfPig.Rendering.Skia), [PdfToSvg.NET](https://github.com/dmester/pdftosvg.net)

### License

DocSharp is licensed under MIT license and can be used for both open source and commercial projects.  

DocSharp.ImageSharp and DocSharp.MagickNET are licensed under [Apache 2.0 license](https://www.apache.org/licenses/LICENSE-2.0.txt).  
ImageSharp has a dual license, please visit [their repository](https://github.com/SixLabors/ImageSharp) for more information. 
VectSharp is used under LGPL in this project (GPL packages are not used).  

DocSharp.Renderer is itself licensed under MIT, but depends on QuestPDF which has additional requirements for companies and may require purchasing a license. Please check [their repository](https://github.com/QuestPDF/QuestPDF?tab=readme-ov-file#fair-and-sustainable-license) for information on the Community and Commercial licenses.  

### Contribute

- If you know how to fix a bug, feel free to open a Pull Request.  
- To implement a new feature, please open an issue or discussion to propose it before working on the pull request.   
- If you find the library useful, adding a star is highly appreciated. Stars are a way to guide other developers towards helpful libraries and tools.
- This is a hobby project. You are welcome to donate to financially support its further development, if you wish (sponsor links for GitHub, LiberaPay, Ko-Fi, BuyMeACoffee and Thanks.dev are available in the repo page).  
