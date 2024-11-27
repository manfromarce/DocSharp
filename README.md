# DocSharp

DocSharp is a pure C# library to convert between document formats without Office interop or native dependencies.

The following packages are currently available:

- DocSharp.Binary: convert Office 97-2003 binary documents (doc, xls, ppt) to OpenXML documents (docx, xlsx, pptx). This is a fork of the abandoned b2xtranslator project which provides critical fixes.
- DocSharp.Docx: convert DOCX to Markdown (and possibly others in the future).

### Supported features

- Binary formats: almost all features are supported, but bugs and exceptions may occur for complex documents. The most obvious and frequent issues from the original project have been fixed, as they were mostly related to using .NET rather than .NET Framework, in particular code pages-based encodings and closed stream / null reference exception for PowerPoint presentations. A wider range of DOC / XLS / PPT documents should now be converted properly, but there are still issues for specific documents I tested. More work is needed to make this library reliable.
- DOCX to Markdown:
  - Text
  - Basic formatting
    - Bold, italic, underline, strikethrough
    - Any highlight color (except none) is converted to `<mark>`
    - H1-H6 headings
  - Inline images
  - Bulleted and numbered lists
  - Paragraph left indent
  - Tables (values only)
  - Hyperlinks
    - Bookmarks are converted to anchors if possible
  - Page and section breaks are converted to horizontal line
  - TODO: math formulas, charts

### Roadmap

- Improve existing packages

- Implement OpenXML renderer, which can be useful for various conversions (Office-specific features can be rasterized or drawn as SVG when converting to simple or older formats).
