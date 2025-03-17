# Supported features (per format)

## DOCX converters

| Feature       | ➡️ RTF | ➡️ Markdown | ➡️ TXT
| ------------- | :-----: | :---------: | :-----: |
| Character formatting | ✅ | ✅ | - |
| Paragraph formatting | ✅ | - | - |
| Section formatting   | ✅ | - | - |
| Page setup   | ✅ | - | - |
| Pictures [1] | ✅ | ✅ | - |
| Lists    | 🟡 | 🟡 | 🟡 |
| Tables   | ✅[3] | 🟡 | 🟡 |
| Styles   | 🟡 [2] | 🟡 [2] | - |
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
| Math | ❌ | ❌ | ❌ |
| OLE objects | ❌ | - | - |
| Form fields | ❌ | - | - |
| Comments | ❌ | - | - |
| Document information | ❌ | - | - |

[1]: RTF supports JPEG, PNG, WMF and EMF; web browsers support JPEG, PNG, GIF, SVG and partially BMP and ICO.  
There is currently no automatic conversion for other image types which may be contained in DOCX.  
The ImagesOutputFolder has to be set to an existing directory, otherwise images are ignored.  
[2]: Preserved as inline formatting  
[3]: Nested tables, conditional formatting and few other advanced properties are not supported yet.

## Markdown renderers

| Feature       | ➡️ DOCX | 
| ------------- | :-----: |
| Emphasis, bold, strikethrough | ✅ |
| Paragraphs | ✅ |
| Headings | ✅ |
| Emojis | ✅ |
| Inline code blocks [1] | ✅ |
| Fenced code blocks [1] | ✅ |
| Quote blocks [1] | ✅ |
| Lists [1] | ✅ |
| Breaks | ✅ |
| Pictures [2] | ✅ |
| Links | ✅ |
| Bookmarks on headings | ✅ |
| HTML inline tags | 🟡 |
| HTML entities | ✅ |
| Raw HTML blocks | ❌ |
| Tables | ✅ |
| Math | ❌ |
| Figures | ❌ |
| Footer | ❌ |
| Task list | ❌ |
| Definition list | ❌ |

[1]: Known issue: lists, code blocks and quotes may not use the correct style when nested inside another quote block.  
[2]: WebP, AVIF and JXL images are not supported in DOCX.  
The ImagesBaseUri property has to be set in order to support relative URLs, otherwise only images referenced with absolute URLs are processed (see the wiki or sample app for examples).
