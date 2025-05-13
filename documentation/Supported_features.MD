# Supported features (per format)

## DOCX converters

| Feature       | ➡️ RTF | ➡️ Markdown | ➡️ HTML | ➡️ TXT |
| ------------- | :-----: | :---------: | :-----: |
| Character formatting | ✅ | ✅ | ✅ | - |
| Paragraph formatting | ✅ | - | ✅ | - |
| Section formatting   | ✅ | - | ❌ | - |
| Page setup   | ✅ | - | ❌ | - |
| Columns | ✅ | ❌ | ✅ | ❌ |
| Pictures [1] | ✅ | ✅ | ✅ | - |
| Lists [2]    | 🟡 | ✅ | ⏳ | ✅ |
| Tables [3]  | ✅ | ✅ | ✅ | ✅ |
| Styles [4] | 🟡 | 🟡 | 🟡 | - |
| Header and footer | ✅ | ❌ | ⏳ | ❌ |
| Endnotes and footnotes | ✅ | ❌ | ⏳ | ❌ |
| Hyperlinks | ✅ | ✅ | ✅ | - |
| Bookmarks | ✅ | ✅ | ✅ | - |
| Drop cap | ✅ | ❌ | ⏳ | ❌ |
| Symbols | ✅ | ✅[7] | ✅[7] | ✅[7] |
| Fields and standard parts [5] | 🟡 | 🟡 | 🟡 | 🟡 |
| Page numbers | ✅ | ❌ | ❌ | ❌ |
| Ruby | ❌ | ❌ | ⏳ | ❌ |
| Shapes | ❌ | ❌ | ❌ | Text is extracted |
| Text boxes | ❌ | ❌ | ❌ | Text is extracted |
| Charts | ❌ | ❌ | ❌ | - |
| Diagrams | ❌ | ❌ | ❌ | - |
| Advanced text effects (includes WordArt) | ❌ | ❌ | 🟡 [6] | - |
| Image effects | ❌ | ❌ | ❌ | - |
| Math | ❌ | ✅ | ⏳ | ❌ |
| OLE objects | ❌ | - | - | - |
| Form fields | ❌ | - | ❌ | - |
| Comments | ❌ | - | ❌ | - |
| Document information | ❌ | - | ❌ | - |

[1]: RTF supports JPEG, PNG, WMF and EMF; web browsers support JPEG, PNG, GIF, SVG and partially BMP and ICO.  
Other image formats can be converted using the DocSharp.ImageSharp or DocSharp.SystemDrawing packages or setting the ImageConverter on a custom converter which returns PNG bytes.  
For Markdown, the ImagesOutputFolder has to be set to an existing directory, otherwise images are ignored.  
For HTML, the ImagesOutputFolder has to be set if desired, otherwise images are preserved as base64.  
See the [Wiki](https://github.com/manfromarce/DocSharp/wiki) for additional information on preserving images in documents.

[2]: There are some known issues with lists, due to the fact that other formats support a subset of DOCX features, and figuring out the correct visual representation is not always easy.

[3]: Nested tables, conditional formatting and few other advanced properties are not supported yet for DOCX --> RTF; while for DOCX --> MD / TXT only basic cell content can be preserved.

[4]: Styles are converted to inline formatting.

[5]: Only the field result is preserved. The field code is preserved in RTF only.  

[6]: Shadow, outline and fill are partially supported; other effects are not supported.  

[7]: Special fonts such as Wingdings, Webdings and Symbol are converted to Unicode characters.  

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
| Styles | 🟡[6] | 🟡[6] | - | ⏳ |
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

[6] preserved as inline formatting if possible.  

## Markdown renderers

| Feature       | ➡️ DOCX | ➡️ RTF |
| ------------- | :-----: | :-----: |
| Emphasis, bold, strikethrough | ✅ | ✅ |
| Paragraphs | ✅ | ✅ |
| Headings | ✅ | ✅ |
| Emojis | ✅ | ✅ |
| Inline code blocks [7] | ✅ | ✅ |
| Fenced code blocks [7] | ✅ | ✅ |
| Syntax highlight for code blocks | ❌ | ❌ |
| Quote blocks [7] | ✅ | ✅ |
| Lists [7] | ✅ | ✅ |
| Task lists | ✅ | ✅ |
| Breaks | ✅ | ✅ |
| Pictures [8] | ✅ | ✅ |
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

[7]: Known issue: code blocks, lists and quotes may not use the correct style when nested inside another quote block.  

[8]: WebP, AVIF and JXL images are not supported in DOCX and RTF, and GIF is also not supported in RTF.  
JPEG and PNG have the broadest support. To convert other images, you can use the DocSharp.ImageSharp or DocSharp.SystemDrawing package or set the ImageConverter property on a custom converter which returns PNG bytes.  
The ImagesBaseUri property has to be set in order to support relative URLs, otherwise only images referenced with absolute URLs are processed (see the wiki or sample app for examples).  
See the [Wiki](https://github.com/manfromarce/DocSharp/wiki) for additional information on preserving images in documents.