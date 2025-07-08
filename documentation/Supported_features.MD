# Supported features (per format)

## DOCX converters

| Feature       | ➡️ RTF | ➡️ HTML | ➡️ Markdown | ➡️ TXT | ➡️ PDF |
| ------------- | :----: | :-----: | :--------: | :----: | :----: |
| Character formatting | ✅ | ✅ | ✅ | - | ⏳ |
| Paragraph formatting | ✅ | ✅ | - | - | ⏳ |
| Section formatting   | ✅ | 🟡 | - | - | ⏳ |
| Page setup   | ✅ | 🟡[9] | - | - | ⏳ |
| Columns | ✅ | ✅ | ❌ | ❌ |⏳ |
| Pictures [1] | ✅ | ✅ | ✅ | - | ⏳ |
| Lists | ✅ | ❌ | ✅[2] | ✅[2] | ⏳ |
| Tables [3]  | ✅ | ✅ | 🟡 | 🟡 | ⏳ |
| Styles [4] | 🟡 | 🟡 | 🟡 | - | ⏳ |
| Header and footer | ✅ | ✅ | ❌ | ❌ | ⏳ |
| Endnotes and footnotes | ✅ | ❌ | ❌ | ❌ | ⏳ |
| Hyperlinks | ✅ | ✅ | ✅ | - | ⏳ |
| Bookmarks | ✅ | ✅ | ✅ | - | ⏳ |
| Drop cap and frame properties | ✅ | ❌ | ❌ | ❌ | ⏳ |
| Symbols | ✅ | ✅[5] | ✅[5] | ✅[5] | ⏳ |
| Page numbers | ✅ | ❌ | ❌ | ❌ | ⏳ |
| Ruby | ❌ | ❌ | ❌ | ❌ | ❌ |
| Modern text effects / WordArt | ❌ | 🟡[9] | ❌ | 🟡 [10] | ❌ |
| Legacy font shadow/outline/emboss/imprint | ✅ | 🟡[9] | ❌ | - | ❌ |
| Modern shapes, text boxes and image/shape effects (DrawingML) | ❌ | ❌ | ❌ | 🟡 [10] | ❌ |
| Legacy shapes, text boxes, effects and WordArt (VML) | ❌ | ❌ | ❌ | 🟡 [10] | ❌ |
| Charts | ❌ | ❌ | ❌ | - | ❌ |
| Diagrams | ❌ | ❌ | ❌ | - | ❌ |
| Math | ✅ | 🟡 | ✅ | ❌ | ⏳ |
| OLE objects [6] | 🟡 | 🟡 | 🟡 | - | ⏳ |
| Fields and standard parts [7] | ✅ | 🟡 | 🟡 | 🟡 | ⏳ |
| Form fields | 🟡[8] | ❌ | - | - | ⏳ |
| Content part (SVG, MathML, SMIL, ...) | ❌ | ✅ SVG, MathML, SMIL | ❌ | - | ⏳ |
| Comments | ❌ | ❌ | - | - | ⏳ |
| Revisions | ❌ | ❌ | - | - | ⏳ |
| Mail merge | ❌ | - | - | - | ❌ |
| Macros | ❌ | - | - | - | ❌ |
| Document information (title, author, ...) | ✅ | ❌ | - | - | ⏳ |

[1]: Only inline images are supported. Floating and wrap-positioned images depend on DrawingML/VML shapes which are not implemented yet.  
RTF supports JPEG, PNG, WMF and EMF; web browsers support JPEG, PNG, GIF, SVG and partially BMP and ICO.  
Other image formats can be converted using the DocSharp.ImageSharp or DocSharp.SystemDrawing packages or setting the ImageConverter on a custom converter which returns PNG bytes.  
For Markdown, the ImagesOutputFolder has to be set to an existing directory, otherwise images are ignored.  
See the [Wiki](https://github.com/manfromarce/DocSharp/wiki) for additional information on preserving images in documents.

[2]: Very limited support for number format and indentation in Markdown and TXT.  

[3]: Nested tables and conditional formatting are not supported yet for DOCX --> RTF; while for DOCX --> MD / TXT only basic cell content can be preserved.

[4]: Styles are converted to inline formatting  

[5]: Special fonts such as Wingdings, Webdings and Symbol are converted to Unicode characters.  

[6]: OLE objects are preserved as images when possible.

[7]: Fields (including table of contents) are supported in RTF; in other formats only the field result/text can be preserved.  
Standard parts are converted to regular (static) content (same behavior as Microsoft Word).

[8]: Modern form fields are not supported in RTF but the result/text is preserved (same behavior as Microsoft Word);  
legacy form fields using field code/char are supported;  
ActiveX controls are preserved as images (like OLE objects).

[9]: Shadow, emboss/imprint, outline and fill effects are partially supported; floating text boxes are not supported.

[10]: Text is extracted from shapes, text boxes and WordArt when converting to TXT but only for inline objects,  
floating objects are ignored to prevent layout issues.

## RTF converters

| Feature       | ➡️ HTML | ➡️ Markdown | ➡️ TXT | ➡️ DOCX |
| ------------- | :------: | :-----: | :---------: | :----: |
| Character formatting | ✅ | ✅ | - | ⏳ |
| Paragraph formatting | ✅ | - | - | ⏳ |
| Section formatting  | 🟡 | - | - | ⏳ |
| Page setup   | ❌ | - | - | ⏳ |
| Pictures | ✅ | ✅ | - | ⏳ |
| External pictures | ✅ | ✅ | - | ⏳ |
| Lists | 🟡 | 🟡 | 🟡 | ⏳ |
| Tables | ✅ | 🟡 | 🟡 | ⏳ |
| Styles | 🟡[11] | 🟡[11] | - | ⏳ |
| Hyperlinks | ✅ | ✅ | - | ⏳ |
| Bookmarks | ✅ | ✅ | - | ⏳ |
| Header and footer | 🟡 | 🟡 | 🟡 | ⏳ |
| Foonotes and endnotes | 🟡 | 🟡 | 🟡 | ⏳ |
| Fields | 🟡 | 🟡 | 🟡 | ⏳ |
| Page numbers | 🟡 | 🟡 | 🟡 | ⏳ |
| Shapes | ❌ | ❌ | ❌ | ⏳ |
| WordArt | ❌ | ❌ | ❌ | ⏳ |
| Image effects | ❌ | ❌ | ❌ | ⏳ |
| Math | ❌ | ❌ | ❌ | ⏳ |
| OLE objects | - | - | - | ⏳ |
| Comments | - | - | - | ⏳ |
| Document information | 🟡 | - | - | ⏳ |

[11] preserved as inline formatting if possible.  

## Markdown renderers

| Feature       | ➡️ DOCX | ➡️ RTF |
| ------------- | :-----: | :-----: |
| Emphasis, bold, strikethrough | ✅ | ✅ |
| Paragraphs | ✅ | ✅ |
| Headings | ✅ | ✅ |
| Emojis | ✅ | ✅ |
| Inline code blocks [12] | ✅ | ✅ |
| Fenced code blocks [12] | ✅ | ✅ |
| Syntax highlight for code blocks | ❌ | ❌ |
| Quote blocks [12] | ✅ | ✅ |
| Lists [12] | ✅ | ✅ |
| Task lists | ✅ | ✅ |
| Breaks | ✅ | ✅ |
| Pictures [13] | ✅ | ✅ |
| Links | ✅ | ✅ |
| Bookmarks on headings | ✅ | ✅ |
| HTML inline tags | 🟡 | ❌ |
| HTML entities | ✅ | ✅ |
| Raw HTML blocks | ❌ | ❌ |
| Tables | ✅ | ✅ |
| Footer and footnotes | ✅ | ✅ |
| Math (LaTex blocks) | ❌ | ❌ |
| Definition lists | ❌ | ❌ |
| Figures | ❌ | ❌ |

[12]: Known issue: code blocks, lists and quotes may not use the correct style when nested inside another quote block.  

[13]: WebP, AVIF and JXL images are not supported in DOCX and RTF, and GIF is also not supported in RTF.  
JPEG and PNG have the broadest support. To convert other images, you can use the DocSharp.ImageSharp or DocSharp.SystemDrawing package or set the ImageConverter property on a custom converter which returns PNG bytes.  
The ImagesBaseUri property has to be set in order to support relative URLs, otherwise only images referenced with absolute URLs are processed (see the wiki or sample app for examples).  
See the [Wiki](https://github.com/manfromarce/DocSharp/wiki) for additional information on preserving images in documents.
