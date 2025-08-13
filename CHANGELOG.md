## 0.12.1 - 2025.08.14

- Support table cell spacing in DOCX to HTML converter and improve tables logic

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

