# Supported features (per format)

## DOCX converters

| Feature       | ➡️ RTF | ➡️ Markdown | ➡️ TXT
| ------------- | :-----: | :---------: | :-----: |
| Character formatting | ✅ | ✅ | - |
| Paragraph formatting | ✅ | - | - |
| Section formatting   | ✅ | - | - |
| Page setup   | ✅ | - | - |
| Columns | ✅ | ❌ | ❌ |
| Pictures [1] | ✅ | ✅ | - |
| Lists [2]    | 🟡 | 🟡 | ✅ |
| Tables [3]  | ✅ | 🟡 | ✅ |
| Styles [4] | 🟡 | 🟡 | - |
| Header and footer | ✅ | ❌ | ❌ |
| Endnotes and footnotes | ✅ | ❌ | ❌ |
| Hyperlinks | ✅ | ✅ | - |
| Bookmarks | ✅ | ❌ | - |
| Drop cap | ✅ | ❌ | ❌ |
| Symbols | ✅ | ✅ | ✅ |
| Fields | 🟡 | ❌ | ❌ |
| Page numbers | ✅ | ❌ | ❌ |
| Ruby | ❌ | ❌ | ❌ |
| Shapes | ❌ | ❌ | Text is extracted |
| Charts | ❌ | ❌ | - |
| Diagrams | ❌ | ❌ | - |
| Advanced text effects (includes WordArt) | ❌ | ❌ | - |
| Image effects | ❌ | ❌ | - |
| Math | ❌ | ✅ | ❌ |
| OLE objects | ❌ | - | - |
| Form fields | ❌ | - | - |
| Comments | ❌ | - | - |
| Document information | ❌ | - | - |

[1]: RTF supports JPEG, PNG, WMF and EMF; web browsers support JPEG, PNG, GIF, SVG and partially BMP and ICO.  
Other image formats can be converted using the DocSharp.ImageSharp or DocSharp.SystemDrawing packages or setting the ImageConverter on a custom converter which returns PNG bytes.  
For Markdown, the ImagesOutputFolder has to be set to an existing directory, otherwise images are ignored.  
See the [Wiki](https://github.com/manfromarce/DocSharp/wiki) for additional information on preserving images in documents.

[2]: There are some known issues with lists, due to the fact that other formats support a subset of DOCX features, and figuring out the correct visual representation is not always easy.

[3]: Nested tables, conditional formatting and few other advanced properties are not supported yet for DOCX --> RTF; while for DOCX --> MD / TXT only basic cell content can be preserved.

[4]: Styles are converted to inline formatting  

## Markdown renderers

| Feature       | ➡️ DOCX | 
| ------------- | :-----: |
| Emphasis, bold, strikethrough | ✅ |
| Paragraphs | ✅ |
| Headings | ✅ |
| Emojis | ✅ |
| Inline code blocks [1] | ✅ |
| Fenced code blocks [1] | ✅ |
| Syntax highlight for code | ❌ |
| Quote blocks [1] | ✅ |
| Lists [1] | ✅ |
| Task lists | ✅ |
| Breaks | ✅ |
| Pictures [2] | ✅ |
| Links | ✅ |
| Bookmarks on headings | ✅ |
| HTML inline tags | 🟡 |
| HTML entities | ✅ |
| Raw HTML blocks | ❌ |
| Tables | ✅ |
| Math (LaTex blocks) | ❌ |
| Figures | ❌ |
| Footer | ❌ |
| Definition list | ❌ |

[1]: Known issue: code blocks, lists and quotes may not use the correct style when nested inside another quote block.  

[2]: WebP, AVIF and JXL images are not supported in DOCX.  
WebP images can be converted using the DocSharp.ImageSharp package or setting the ImageConverter property on a custom converter which returns PNG bytes.  
The ImagesBaseUri property has to be set in order to support relative URLs, otherwise only images referenced with absolute URLs are processed (see the wiki or sample app for examples).  
See the [Wiki](https://github.com/manfromarce/DocSharp/wiki) for additional information on preserving images in documents.
