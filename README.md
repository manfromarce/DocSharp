# DocSharp

DocSharp is a pure C# library to convert between document formats without Office interop or native dependencies.

The following packages are currently available:

- DocSharp.Binary: convert Office 97-2003 binary documents (doc, xls, ppt) to OpenXML documents (docx, xlsx, pptx). This is a fork of the abandoned [b2xtranslator project](https://github.com/EvolutionJobs/b2xtranslator) which provides critical fixes.
- DocSharp.Docx: convert DOCX to RTF, HTML, Markdown or LaTex
- DocSharp.Rtf: convert RTF to DOCX, HTML or Markdown.
- DocSharp.Markdig: convert Markdown to DOCX or RTF

There is no common DOM to manipulate documents or generate them from scatch, this library is for conversion only.

To manipulate Office Open XML documents, you can use the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) or other libraries built on top of it: [OfficeIMO](https://github.com/EvotecIT/OfficeIMO), [OpenXML-Office](https://github.com/DraviaVemal/OpenXML-Office), [ClosedXML](https://github.com/ClosedXML/ClosedXML), [ShapeCrawler](https://github.com/ShapeCrawler/ShapeCrawler).

### Supported features

- Binary formats: almost all doc/xls/ppt features were supported by the original project, but exceptions occurred when using .NET (rather than .NET Framework) or loading specific documents. Most errors should be fixed now but more work is needed to make the library reliable; if you find other bugs, you are welcome to open an issue (please attach a sample file if the issue only occurs for specific documents).
- DOCX to Markdown:
  * Text and basic formatting
    - Bold, italic, underline, strikethrough
    - Any highlight color (except none) is converted to `<mark>`
  * Inline images
    - `ImagesOutputFolder` needs to be set to an existing directory, otherwise images are skipped. An absolute URI is used by default; to produce a relative URI set `ImagesBaseUriOverride` to any not-null folder path (empty string or "." means same folder as the Markdown file, "../images" means images subfolder in the parent folder).
    - Only `Pict` elements are currently recognized, other image types are not implemented (e.g. WordPad embeds images in a different way).
  * Tables (values only)
  * External hyperlinks
  * Page breaks are converted to horizontal lines
  * TODO: styles (including H1-H6 headers), bookmarks (internal hyperlinks), lists, special chars, math formulas, charts

### Roadmap

- Publish NuGet packages
- Support more elements and attributes
- Documentation (for now you can refer to the sample app). When ready, any documentation will be available in the project Wiki.

### Credits

Dependencies: 
- [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)
- [Markdig](https://github.com/xoofx/markdig)

Forked: 
- [b2xtranslator](https://github.com/EvolutionJobs/b2xtranslator)
