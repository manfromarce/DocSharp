# DocSharp

DocSharp is a pure C# library to convert between document formats without Office interop or native dependencies.

The following packages are currently available:

- DocSharp.Binary: convert Office 97-2003 binary documents (doc, xls, ppt) to OpenXML documents (docx, xlsx, pptx).
- DocSharp.Docx: convert DOCX to Markdown (and possibly others in the future).
- DocSharp.Renderer: render OpenXML documents (currently DOCX only) to PDF, images or SVG using VectSharp.
