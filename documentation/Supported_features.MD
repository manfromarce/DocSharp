# Supported formats (direct conversion)

| From \ To | DOCX | XLSX| PPTX | RTF | HTML | Markdown | TXT | Renderer (PDF, XPS, SVG, Images) |
| --------- | :--: | :--: | :--: | :--: | :----: | :-: | :------------------------------: |
| DOCX      |  -   |  -  |  -   |  ✅  |  ✅  |  ✅  |  ✅ |               ⏳                |
| DOC       |  ✅  |  -  |  -  |  ❌  |  ❌  |  ❌  |  ❌ |               ❌                 |
| XLS       |  -   |  ✅ |  -  |  ❌  |  ❌  |  ❌  |  ❌ |              ❌                 |
| PPT       |  -   |  -  |  ✅  |  ❌ |  ❌  |  ❌  |  ❌ |              ❌                 |
| RTF       |  Planned  | - | - |  -  |  ❌  |  ❌  |  ❌  |               ❌                |
| ODT       |  Planned  | - | - |  ❌  |  ❌ |  ❌  |  ❌  |               ❌                |
| Markdown  |  ✅  |  ❌  |  ❌  | ❌ | ❌ |   -   | ❌ |           ❌                      |

Other conversions (not planned / out of scope):  
- For RTF to HTML, you can convert RTF -> DOCX -> HTML, or use [RtfPipe](https://github.com/erdomke/RtfPipe)
- For HTML to DOCX, use [HtmlToOpenXml](https://github.com/onizet/html2openxml)
- For HTML to Markdown, use [ReverseMarkdown](https://github.com/mysticmind/reversemarkdown-net)
- For HTML to PDF, you can use [HTMLToQPDF](https://github.com/Relorer/HTMLToQPDF) or [PeachPdf](https://github.com/jhaygood86/PeachPDF) for simple HTML, or libraries such as [Puppeteer Sharp](https://github.com/hardkoded/puppeteer-sharp) or similar for complex web pages (these depend on Chromium).
- For Markdown to HTML, use [Markdig](https://github.com/xoofx/markdig)
- For Markdown to PDF rendering, use VectSharp or [QuestPDF.Markdown](https://github.com/christiaanderidder/QuestPDF.Markdown)

# Supported features (per format)

## DOCX converters

| Feature       | ➡️ RTF | ➡️ HTML | ➡️ Markdown | ➡️ TXT |
| ------------- | :----: | :-----: | :--------: | :----: |
| Character formatting | ✅ | ✅ | ✅ | - |
| Paragraph formatting | ✅ | ✅ | - | - |
| Section formatting   | ✅ | 🟡 | - | - |
| Page setup   | ✅ | - | - | - |
| Columns | ✅ | ✅ | ❌ | ❌ 
| Pictures [1] | ✅ | ✅ | ✅ | - |
| Lists | ✅ | ✅ | ✅[2] | ✅[2] |
| Tables [3]  | ✅ | ✅ | 🟡 | 🟡 |
| Styles [4] | 🟡 | 🟡 | 🟡 | - |
| Header and footer | ✅ | ✅ | ✅ | ✅ |
| Footnotes and endnotes | ✅ | ✅ | ✅ | ✅ |
| Hyperlinks | ✅ | ✅ | ✅ | - |
| Bookmarks | ✅ | ✅ | ✅ | - |
| Drop cap and frame properties | ✅ | ❌ | ❌ | ❌ |
| Symbols | ✅ | ✅[5] | ✅[5] | ✅[5] |
| Page numbers | ✅ | ❌ | ❌ | ❌ |
| Ruby | ❌ | ✅ | ❌ | ❌ |
| Modern text effects and WordArt | -[11] | 🟡[9] | ❌ | 🟡 [10] |
| Legacy font shadow/outline/emboss/imprint | ✅ | 🟡[9] | ❌ | - |
| Modern shapes, text boxes and image/shape effects (DrawingML) | ⏳[14] | ❌ | ❌ | 🟡 [10] |
| Legacy shapes, text boxes, effects and WordArt (VML) | ⏳[14] | ❌ | ❌ | 🟡 [10] |
| SmartArt diagrams | ⏳[14] | ❌ | ❌ | - |
| Legacy diagrams | ⏳[14] | ❌ | ❌ | - |
| Charts | -[15] | ❌ | ❌ | - |
| Math | ✅ | 🟡 | ✅ | ❌ |
| Ink [6] | 🟡 | 🟡 | 🟡 | - |
| Signature line [6] | 🟡 | 🟡 | 🟡 | - |
| OLE objects [6] | 🟡 | 🟡 | 🟡 | - |
| 3D; videos [6] | 🟡 | 🟡 | 🟡 | - |
| Fields and standard parts [7] | ✅ | 🟡 | 🟡 | 🟡 |
| Form fields | 🟡[8] | ❌ | ❌ | ❌ |
| Comments | ✅ | ❌ | - | - |
| Revisions | ❌ | ❌ | - | - |
| Mail merge | ❌ | - | - | - |
| Macros | ❌ | - | - | - |
| Document background | 🟡[12] | 🟡[13] | - | - |
| Document information (title, author, ...) | ✅ | ❌ | - | - |
| Document settings and protection | 🟡 | - | - | - |
| Sub-documents / Alternate content | ✅ DOCX, RTF | ✅ DOCX, HTML, SVG, MathML, SMIL | ✅ DOCX | ✅ DOCX |

[1]: Wrap layouts and floating images are preserved in RTF, as well as few basic transforms (crop, flip, rotate).  
Other converters support inline images only, without transforms/effects.  
RTF supports JPEG, PNG, WMF and EMF formats; web browsers support JPEG, PNG, GIF, SVG and partially BMP and ICO.  
Other image formats can be converted using the DocSharp.ImageSharp or DocSharp.SystemDrawing packages or setting the `ImageConverter` on a custom converter which returns PNG bytes.  
For Markdown, the `ImagesOutputFolder` has to be set to an existing directory, otherwise images are ignored.  
See the [Wiki](https://github.com/manfromarce/DocSharp/wiki) for additional information on preserving images in documents.

[2]: Very limited support for number format and indentation in Markdown and TXT.  

[3]: In DOCX --> MD / TXT conversion only basic cell content can be preserved.  
HTML has minor limitations, such as diagonal borders inside cells are not supported.  
RTF should support everything, including conditional formatting and nested tables.

[4]: Styles are converted to inline formatting.  

[5]: Special fonts such as Wingdings, Webdings and Symbol are converted to Unicode characters.  

[6]: Ink, signature line, OLE objects and 3D/media are preserved as images when possible.  
Word processors usually write a picture fallaback for these elements.

[7]: Fields (including table of contents) are supported in RTF; in other formats only the field result/text can be preserved.  
Standard/quick parts are converted to regular static content because they are not supported in RTF (same behavior as Microsoft Word).

[8]: Modern form fields are not supported in RTF but the result/text is preserved (same behavior as Microsoft Word);  
legacy form fields using field code/char are supported;  
ActiveX controls are preserved as images (like OLE objects).

[9]: Basic shadow, emboss/imprint, outline and fill effects are partially supported; floating text boxes are not supported.

[10]: Text is extracted from shapes, text boxes and WordArt when converting to TXT but only for inline objects,  
floating objects are ignored to prevent layout issues.

[11]: Modern text effects are not supported in RTF and removed by Microsoft Word too.

[12]: Most background types are supported in DOCX to RTF conversion, including solid colors, gradients, patterns, textures and pictures.  
Only watermarks are not supported because they require full VML support in order to translate text paths.

[13]: Only solid background is supported in HTML.

[14]: Support for drawings is still incomplete:  
- **Supported**: DrawingML shapes / lines / arrows (created by modern Word versions, unless in compatibility mode), text content, solid fill, gradient fill, picture/texture fill, line color, line style (dashed, dotted, double, cap, ...), line thickness
- **Not supported**: VML (Word 2007 or compatibility mode), manual adjustments (corners / points adjustments and similar), freehand curve/polygon, some shapes that are not available in RTF and would require conversion to custom shapes, pattern fill, gradient/pattern for line, some advanced text layout options, effects, 3D effects, WordArt, diagrams, shape groups, drawing canvas
- **Partially supported**: shapes in line with text (not floating) are preserved but not aligned properly.

[15]: "Real" charts are not available in the RTF specification, so MS Word converts them to images or OLE objects.  
This operation is complex and currently considered out-of-scope for this library.  
However, some charts may already be OLE objects (Excel/MS Graph object; usually happens if the document was originally a Word 2007 DOCX, or a DOC/RTF converted to DOCX keeping compatibility mode), and these are supported. 

## Markdown renderers

| Feature       | ➡️ DOCX | ➡️ RTF |
| ------------- | :-----: | :-----: |
| Emphasis (bold/italic), strikethrough | ✅ | ✅ |
| Paragraphs | ✅ | ✅ |
| Headings | ✅ | ✅ |
| Emojis | ✅ | ✅ |
| Inline code blocks [16] | ✅ | ✅ |
| Fenced code blocks [16] | ✅ | ✅ |
| Syntax highlight for code blocks | ❌ | ❌ |
| Quote blocks [16] | ✅ | ✅ |
| Lists [15] | ✅ | ✅ |
| Task lists | ✅ | ✅ |
| Breaks | ✅ | ✅ |
| Pictures [17] | ✅ | ✅ |
| Links | ✅ | ✅ |
| Bookmarks on headings | ✅ | ✅ |
| HTML inline tags | 🟡[18] | ❌ |
| HTML entities | ✅ | ✅ |
| Raw HTML blocks | ❌ | ❌ |
| Tables | ✅ | ✅ |
| Footer and footnotes | ✅ | ✅ |
| Math (LaTex blocks) | ❌ | ❌ |
| Definition lists | ❌ | ❌ |
| Figures | ❌ | ❌ |
| Diagrams (Mermaid) | ❌ | ❌ |

[16]: Known issue: code blocks, lists and quotes may not use the correct style when nested inside another quote block.  

[17]: WebP, AVIF and JXL images are not supported in DOCX and RTF, and GIF is also not supported in RTF.  
JPEG and PNG have the broadest support. To convert other images, you can use the DocSharp.ImageSharp or DocSharp.SystemDrawing package or set the `ImageConverter` property on a custom converter which returns PNG bytes.  
The `ImagesBaseUri` property has to be set in order to support relative URLs, otherwise only images referenced with absolute URLs are processed (see the wiki or sample app for examples).  
See the [Wiki](https://github.com/manfromarce/DocSharp/wiki) for additional information on preserving images in documents.

[18]: Only these basic tags are supported: `u`, `i`, `b`, `em`, `strong`, `sub`, `sup`, `mark`, `s`, `strike`, `del`, `ins`
