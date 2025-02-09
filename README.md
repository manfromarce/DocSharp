# DocSharp

DocSharp is a pure C# library to convert between document formats without Office interop or native dependencies.

The following packages are currently available:

- DocSharp.Binary: convert Office 97-2003 binary documents (doc, xls, ppt) to OpenXML documents (docx, xlsx, pptx). This is a fork of the abandoned [b2xtranslator project](https://github.com/EvolutionJobs/b2xtranslator) which provides critical fixes. 
- DocSharp.Docx: convert DOCX to Markdown and RTF. Possible applications include generating Open XML documents in C# and exporting for other editors, or loading Word documents in a RichTextBox / RichEditBox control.
- DocSharp.Markdown: convert Markdown to DOCX using a custom Markdig renderer.

Packages can be installed via NuGet:  
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Doc?style=flat-square&label=DocSharp.Binary.Doc)](https://www.nuget.org/packages/DocSharp.Binary.Doc/) 
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Xls?style=flat-square&label=DocSharp.Binary.Xls)](https://www.nuget.org/packages/DocSharp.Binary.Xls/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Binary.Ppt?style=flat-square&label=DocSharp.Binary.Ppt)](https://www.nuget.org/packages/DocSharp.Binary.Ppt/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Docx?style=flat-square&label=DocSharp.Docx)](https://www.nuget.org/packages/DocSharp.Docx/)
[![NuGet](https://img.shields.io/nuget/vpre/DocSharp.Markdown?style=flat-square&label=DocSharp.Markdown)](https://www.nuget.org/packages/DocSharp.Markdown/)

There is no common DOM to manipulate or generate documents, this library is mainly for conversion. However, the Docx package provides some helper methods on top of the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) that may be extended in the future.  
If your main purpose is creating documents from scratch you can consider the following libraries: [OfficeIMO](https://github.com/EvotecIT/OfficeIMO), [OpenXML-Office](https://github.com/DraviaVemal/OpenXML-Office), [ClosedXML](https://github.com/ClosedXML/ClosedXML), [ShapeCrawler](https://github.com/ShapeCrawler/ShapeCrawler), [QuestPDF](https://github.com/QuestPDF/QuestPDF).

### Supported features

- Binary formats: almost all doc/xls/ppt features were supported by the original project, but exceptions occurred when using .NET (rather than .NET Framework) or loading specific documents/encodings. Most errors should be fixed now but more work is needed to make the library reliable; if you find other bugs, you are welcome to open an issue (please attach a sample file if the issue only occurs for specific documents).
- DOCX to RTF: 
  * Text and most font formatting
  * Paragraph options, lists and tables - not all properties are supported, but should be sufficient for medium documents.
  * Images:
    - JPEG, PNG and EMF are supported. 
    - Only inline images are supported (wrap layouts are not yet implemented).
  * Hyperlinks and bookmarks
  * Page setup: size, orientation, margins, borders
  * Header and footer
  * TODO: math formulas, drawings, OLE objects, fields and more
- DOCX to Markdown:
  * Text and basic formatting
    - Bold, italic, underline, strikethrough, superscript, subscript
    - Any highlight color is converted to `<mark>`
  * Inline images
    - `ImagesOutputFolder` needs to be set to an existing directory, otherwise images are skipped. An absolute URI is used by default; to produce a relative URI set `ImagesBaseUriOverride` to any not-null folder path (empty string or "." means same folder as the Markdown file, "../images" means images subfolder in the parent folder).
    - Some image types are not recognized (e.g. WordPad embeds images in a different way compared to MS Word and other word processors).
    - Images should be in JPEG, PNG or GIF format to be supported by browsers; BMP is partially supported but not recommended. There is currently no automatic image conversion implemented.
    - Crop and effects are not supported.
  * Tables (values only)
  * External hyperlinks
  * Page breaks are converted to horizontal lines
  * TODO: H1-H6 headers (Word styles), bookmarks (internal hyperlinks), lists, math formulas, charts; support for encrypted Word documents
- Markdown to DOCX:
  * Basic Markdown features
  * External hyperlinks
  * Bookmarks for internal hyperlinks to headings (GitHub-like auto-identifiers)
  * Images
    - The converter attempts to read local images and download online images (http/https URLs only). If this behavior is not desired, set `SkipImages` to true.
    - Images specified as absolute URLs are processed by default. For relative URLs `ImagesBaseUri` needs to be set to an absolute local directory path or http(s) URL, which will be combined with the image URL at runtime, such as: `C:\Data` + `./images/image1.jpg` (all kind of URIs should be recognized).
    - WEBP and AVIF images are ignored as they are not supported in DOCX documents; base64 is also ignored as it is rarely used and not supported by many Markdown processors.
  * Tables (experimental)
  * TODO: other internal hyperlinks types, HTML tags (`<u>`, `<sup>`, `<sub>`, `<mark>`, ...), math and other extensions

### Roadmap

- Support more elements and attributes, and fix issues on edge cases
- Reverse RTF to DOCX conversion
- Documentation: for now you can refer to the sample app. When ready, any documentation will be available in the project [Wiki](https://github.com/manfromarce/DocSharp/wiki).

### Credits

Dependencies: 
- [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)
- [Markdig](https://github.com/xoofx/markdig) - for DocSharp.Markdown only

Forked: 
- [b2xtranslator](https://github.com/EvolutionJobs/b2xtranslator)
- [markdig.docx](https://github.com/morincer/markdig.docx)

Others:
- [Html2OpenXml](https://github.com/onizet/html2openxml) for images header decoding and unit conversions.

### License

DocSharp is licensed under MIT license and can be used for both open source and commercial projects.  
If you find the library useful, adding a star is highly appreciated.
