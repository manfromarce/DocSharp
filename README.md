# DocSharp

DocSharp is a pure C# library to convert between document formats without Office interop or native dependencies.

The following packages are currently available:

- DocSharp.Binary: convert Office 97-2003 binary documents (doc, xls, ppt) to OpenXML documents (docx, xlsx, pptx). This is a fork of the abandoned [b2xtranslator project](https://github.com/EvolutionJobs/b2xtranslator) which provides critical fixes.
- DocSharp.Docx: convert DOCX to Markdown (and possibly others in the future).

### Supported features

- Binary formats: almost all features are supported, but bugs and exceptions may occur for complex documents. The most obvious and frequent issues from the original project have been fixed, as they were mostly related to using .NET rather than .NET Framework, in particular code pages-based encodings and closed stream / null reference exception for PowerPoint presentations. A wider range of DOC / XLS / PPT documents should now be converted properly, but there are still issues for specific documents I tested. More work is needed to make this library reliable.
- DOCX to Markdown:
  - Text and basic formatting
    - Bold, italic, underline, strikethrough
    - Any highlight color (except none) is converted to `<mark>`
  - Inline images
    - `ImagesOutputFolder` needs to be set to an existing directory. An absolute URI is added by default; to produce a relative URI set `ImagesBaseUriOverride` to ".", an empty string or any desired relative path.
    - Only `Pict` elements are currently recognized, other image types are not implemented (e.g. WordPad embeds images in a different way).
  - External hyperlinks
  - Tables (values only)
  - Page breaks are converted to horizontal lines
  - TODO: H1-H6 headers, bookmarks (internal hyperlinks), lists, special chars, math formulas, charts

### Roadmap

- Support more elements and attributes.
- Consider other conversions such as RTF to DOCX and DOCX to RTF.
- Implement an OpenXML renderer as a separate package. It can be useful for various conversions, as Office-specific features need to be rasterized or drawn as SVG when converting to other formats.
