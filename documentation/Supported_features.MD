# Supported features (per format)

## DOCX converters

| Feature       | ➡️ RTF | ➡️ Markdown | ➡️ TXT |
| ------------- | :-----: | :---------: | :-----: |
| Character formatting | ✅ | ✅ | - |
| Paragraph formatting | ✅ | - | - |
| Section formatting   | ✅ | - | - |
| Page setup   | ✅ | - | - |
| Columns | ✅ | ❌ | ❌ |
| Pictures [1] | ✅ | ✅ | - |
| Lists [2]    | ✅ | ✅ | ✅ |
| Tables [3]  | ✅ | 🟡 | 🟡 |
| Styles [4] | 🟡 | 🟡 | - |
| Header and footer | ✅ | ❌ | ❌ |
| Endnotes and footnotes | ✅ | ❌ | ❌ |
| Hyperlinks | ✅ | ✅ | - |
| Bookmarks | ✅ | ✅ | - |
| Drop cap | ✅ | ❌ | ❌ |
| Symbols | ✅ | ✅[5] | ✅[5] |
| Page numbers | ✅ | ❌ | ❌ |
| Ruby | ❌ | ❌ | ❌ |
| Shapes and text boxes (DrawingML) | ❌ | ❌ | Text is extracted |
| Shapes and text boxes (VML) | ❌ | ❌ | Text is extracted |
| Charts | ❌ | ❌ | - |
| Diagrams | ❌ | ❌ | - |
| Advanced text effects (includes WordArt) | ❌ | ❌ | - |
| Image effects | ❌ | ❌ | - |
| Math | ✅ | ✅ | ❌ |
| OLE objects | ❌[6] | - | - |
| Fields and standard parts [7] | ✅ | 🟡 | 🟡 |
| Form fields | 🟡[8] | - | - |
| Content part (SVG, MathML, SMIL, ...) | ❌ | ❌ | - |
| Comments | ❌ | - | - |
| Revisions | ❌ | - | - |
| Mail merge | ❌ | - | - |
| Document information (title, author, ...) | ✅ | - | - |

[1]: Only inline images are supported. Floating and wrap-positioned images depend on DrawingML/VML shapes which are not implemented yet.  
RTF supports JPEG, PNG, WMF and EMF; web browsers support JPEG, PNG, GIF, SVG and partially BMP and ICO.  
Other image formats can be converted using the DocSharp.ImageSharp or DocSharp.SystemDrawing packages or setting the ImageConverter on a custom converter which returns PNG bytes.  
For Markdown, the ImagesOutputFolder has to be set to an existing directory, otherwise images are ignored.  
See the [Wiki](https://github.com/manfromarce/DocSharp/wiki) for additional information on preserving images in documents.

[2]: There are some known issues with lists, due to the fact that other formats support a subset of DOCX features, and figuring out the correct visual representation is not always easy.

[3]: Nested tables and conditional formatting are not supported yet for DOCX --> RTF; while for DOCX --> MD / TXT only basic cell content can be preserved.

[4]: Styles are converted to inline formatting  

[5]: Special fonts such as Wingdings, Webdings and Symbol are converted to Unicode characters.  

[6]: OLE objects are preserved as images in most cases when converting to RTF.  

[7]: Fields (including table of contents) are supported in RTF; in other formats only the field result/text can be preserved.  
Standard parts are converted to regular (static) content (same behavior as Microsoft Word).

[8]: Modern form fields are not supported in RTF but the result/text is preserved (same behavior as Microsoft Word);  
legacy form fields using field code/char are supported;  
ActiveX controls are preserved as images (like OLE objects).

## Markdown renderers

| Feature       | ➡️ DOCX | ➡️ RTF |
| ------------- | :-----: | :-----: |
| Emphasis, bold, strikethrough | ✅ | ✅ |
| Paragraphs | ✅ | ✅ |
| Headings | ✅ | ✅ |
| Emojis | ✅ | ✅ |
| Inline code blocks [1] | ✅ | ✅ |
| Fenced code blocks [1] | ✅ | ✅ |
| Syntax highlight for code blocks | ❌ | ❌ |
| Quote blocks [1] | ✅ | ✅ |
| Lists [1] | ✅ | ✅ |
| Task lists | ✅ | ✅ |
| Breaks | ✅ | ✅ |
| Pictures [2] | ✅ | ✅ |
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

[1]: Known issue: code blocks, lists and quotes may not use the correct style when nested inside another quote block.  

[2]: WebP, AVIF and JXL images are not supported in DOCX and RTF, and GIF is also not supported in RTF.  
JPEG and PNG have the broadest support. To convert other images, you can use the DocSharp.ImageSharp or DocSharp.SystemDrawing package or set the ImageConverter property on a custom converter which returns PNG bytes.  
The ImagesBaseUri property has to be set in order to support relative URLs, otherwise only images referenced with absolute URLs are processed (see the wiki or sample app for examples).  
See the [Wiki](https://github.com/manfromarce/DocSharp/wiki) for additional information on preserving images in documents.
