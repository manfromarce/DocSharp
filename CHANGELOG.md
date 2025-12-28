## 0.18.0 - not published yet

- Added conversion options to `WordprocessingDocument.SaveTo` extension methods
- Added ability to append to existing Markdown / plain text when converting DOCX
- Markdown to DOCX / RTF converter: added ability to configure page size and margins
- Improved input/output encodings handling
- Code refactor for maintainability and extensibility
- Remove .NET 6 target, now out of support
- Created GitHub Actions workflow for build, test and NuGet publish
- Other improvements

## 0.17.0 - 2025.12.04

- Add .NET 10 target and update dependencies
- Fix: appending Markdown with links to a DOCX that already contains one (or more) hyperlinks resulted in a corrupted document
- Changed the default DOCX template culture (for DocSharp.Markdown) to en-US

## 0.16.0 - 2025.10.26

- Support shape type and outline color/width for pictures in DOCX to RTF converter  
  (some issues remaining with inline pictures: line color might not be correct and line dash style is not preserved)
- Support custom dash styles for shape outlines in DOCX to RTF converter
- Fix table grid iteration and bounds checks in DOC converter (PR [#14](https://github.com/manfromarce/DocSharp/pull/14))

## 0.15.0 - 2025.09.12

- Support for shapes, lines and text boxes in DOCX to RTF converter.  
  Not everything is supported, see [Supported features](https://github.com/manfromarce/DocSharp/blob/main/documentation/Supported_features.MD) for reference.

## 0.14.0 - 2025.09.03

- Improvements for binary to Open XML converters
- Add "MD Table" style in Markdown to DOCX renderer; this can be used to customize table appearance (like other styles)
- Fix: empty table cells sometimes caused the document to become corrupted in Markdown to DOCX converter
- Fix: don't reset numbering in the target document when appending Markdown to an existing DOCX document (issue #12)
- Improvements to styles and lists handling when appending Markdown to an existing DOCX document (issue #12)
- Heading styles used in Markdown to DOCX converter now inherit from default Word heading styles, so that collapsing/expanding is available.
- Fix: SVG images are now preserved in Markdown to DOCX renderer
- Improve styles logic in DOCX to RTF/HTML/MD converters

## 0.13.1 - 2025.08.31

- Fixed issue with stream support for binary converters (DOC to DOCX, XLS to XLSX, PPT to PPTX)

## 0.13.0 - 2025.08.30

- Support table cell spacing in DOCX to HTML converter
- Preserve table left indentation correctly in DOCX to HTML converter
- Fix issues with margins and borders, and improve overall tables logic in DOCX to RTF/HTML converter
- Support paragraph borders spacing (distance from text) in DOCX to HTML converter
- Recognize and convert color names (e.g. "red") in DOCX to RTF/HTML converter
- Stream support for binary converters (DOC to DOCX, XLS to XLSX, PPT to PPTX)
- Fix issue with line spacing in DOC to DOCX and DOCX to HTML conversion
- Fix issues with table left indentation in DOCX to RTF converter
- Other bug fixes

## 0.12.0 - 2025.08.12

- Fix issues in PPT to PPTX converter and improve performance
- Fix issue with emojis (surrogate pairs chars) in DOC to DOCX converter
- Hidden paragraphs are no longer exported in DOCX to HTML/Markdown/TXT conversion
- Hidden runs are no longer exported in DOCX to Markdown/TXT conversion and are hidden in DOCX to HTML
- Support for inset, outset, groove and ridge border styles in DOCX to HTML converter
- Support for 3D and dash-dot stripe border styles in DOCX to RTF converter
- Try to preserve text fill effect as font color in DOCX to RTF converter
- Preserve "leading zeros" format in numbered lists in DOCX to HTML/TXT converter
- Fix paragraph spacing and automatic cell height for default tables in DOCX to HTML converter
- Fix row height not detected in some cases in DOCX to HTML converter
- Fix issues with vertical text in table cells in DOCX to RTF/HTML converter

**Full Changelog**: https://github.com/manfromarce/DocSharp/compare/v0.11.0...v0.12.0

## 0.11.0

- Preserve most settings.xml elements in DOCX to RTF converter; this also fixes inaccurate paragraph spacing in some cases
- Improve paragraph borders, spacing and indentation handling in DOCX to RTF/HTML converters
- Fix sections when no SectionProperties is present (e.g. document produced by WordPad)
- Fix issues with multi-level lists in DOCX to HTML/TXT converters
- Fix text not wrapping in DOCX to HTML converter
- Enable Dependabot to automate dependency updates
- Other improvements

**Full Changelog**: https://github.com/manfromarce/DocSharp/compare/v0.10.0...v0.11.0

## 0.10.0 - 2025.08.04

- Fix issue with run position
- Fix issues with lists in DOCX documents produced by WordPad
- Fix issue with images and OLE objects in DOCX documents produced by WordPad

**Full Changelog**: https://github.com/manfromarce/DocSharp/compare/v0.9.0...v0.10.0

## 0.9.0 - 2025.08.04

- Fix issue in XLS to XLSX converter
- Add support for width and height generic attributes in Markdown renderer
- Features and improvements for the DOCX converters from the dev branch have been merged, including:
    * New DOCX to HTML converter
    * Support for nested tables and conditional formatting in DOCX to RTF converter
    * Support for wrap layouts and floating images in DOCX to RTF converter
    * Support for comments in DOCX to RTF converter
    * Support for DOCX sub-documents
    * Preserve ink, signature lines, media elements and 3D objects as images
- Other improvements

**Full Changelog**: https://github.com/manfromarce/DocSharp/compare/v0.8.5...v0.9.0

