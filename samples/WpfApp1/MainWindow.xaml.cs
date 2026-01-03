using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.Win32;
using DocSharp.Binary.DocFileFormat;
using DocSharp.Binary.Spreadsheet.XlsFileFormat;
using DocSharp.Binary.PptFileFormat;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Docx;
using DocSharp.Imaging;
using DocSharp.Markdown;
using DocSharp.Renderer;
using HtmlToOpenXml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace WpfApp1;
/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void BinaryToOpenXml_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Multiselect = true,
            Filter = "Office 97-2003 documents|*.doc;*.dot;*.xls;*.xlt;*.xlr;*.ppt;*.pps;*.pot",
        };
        if (ofd.ShowDialog(this) == true)
        {
            var folderDlg = new OpenFolderDialog()
            {
                Multiselect = false,
            };
            if (folderDlg.ShowDialog(this) == true)
            {
                try
                {
                    string outputDir = folderDlg.FolderName;
                    foreach (string file in ofd.FileNames)
                    {
                        string inputExt = Path.GetExtension(file).ToLower();
                        using (var reader = new StructuredStorageReader(file))
                        {
                            string outputExt = inputExt + "x";
                            string baseName = Path.GetFileNameWithoutExtension(file);
                            string outputFile = Path.Join(outputDir, baseName + outputExt);
                            switch (inputExt)
                            {
                                case ".doc":
                                case ".dot":
                                    var doc = new WordDocument(reader);
                                    var docxType = inputExt == ".dot" ? DocSharp.Binary.OpenXmlLib.WordprocessingDocumentType.Template :
                                                                        DocSharp.Binary.OpenXmlLib.WordprocessingDocumentType.Document;
                                    using (var docx = DocSharp.Binary.OpenXmlLib.WordprocessingML.WordprocessingDocument.Create(outputFile, docxType))
                                    {
                                        DocSharp.Binary.WordprocessingMLMapping.Converter.Convert(doc, docx);
                                    }
                                    break;
                                case ".xls":
                                case ".xlt":
                                    var xls = new XlsDocument(reader);
                                    var xlsxType = inputExt == ".xlt" ? DocSharp.Binary.OpenXmlLib.SpreadsheetDocumentType.Template :
                                                                        DocSharp.Binary.OpenXmlLib.SpreadsheetDocumentType.Workbook;
                                    using (var xlsx = DocSharp.Binary.OpenXmlLib.SpreadsheetML.SpreadsheetDocument.Create(outputFile, xlsxType))
                                    {
                                        DocSharp.Binary.SpreadsheetMLMapping.Converter.Convert(xls, xlsx);
                                    }
                                    break;
                                case ".ppt":
                                case ".pps":
                                case ".pot":
                                    var ppt = new PowerpointDocument(reader);
                                    var pptxType = DocSharp.Binary.OpenXmlLib.PresentationDocumentType.Presentation;
                                    if (inputExt == ".pot")
                                    {
                                        pptxType = DocSharp.Binary.OpenXmlLib.PresentationDocumentType.Template;
                                    }
                                    else if (inputExt == ".pps")
                                    {
                                        pptxType = DocSharp.Binary.OpenXmlLib.PresentationDocumentType.Slideshow;
                                    }
                                    using (var pptx = DocSharp.Binary.OpenXmlLib.PresentationML.PresentationDocument.Create(outputFile, pptxType))
                                    {
                                        DocSharp.Binary.PresentationMLMapping.Converter.Convert(ppt, pptx);
                                    }
                                    break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void DocxToRtf_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Word OpenXML document|*.docx;*.dotx;*.docm;*.dotm",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Rich Text Format|*.rtf",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".rtf"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var converter = new DocxToRtfConverter()
                    {
                        ImageConverter = new ImageSharpConverter(), // Converts TIFF, GIF and other formats which are not supported in RTF.
                        OriginalFolderPath = Path.GetDirectoryName(ofd.FileName), // converts sub-documents (if any)
                        OutputFolderPath = Path.GetDirectoryName(sfd.FileName)
                    };
                    converter.Convert(ofd.FileName, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void DocxToHtml_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Word OpenXML document|*.docx;*.dotx;*.docm;*.dotm",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "HTML|*.html;*.htm",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".html"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var converter = new DocxToHtmlConverter()
                    {
                        ExportHeaderFooter = true,
                        ExportFootnotesEndnotes = true,
                        ImageConverter = new SystemDrawingConverter(), // Converts TIFF, WMF and EMF
                                                                       // (ImageSharp does not support WMF / EMF yet)
                        OriginalFolderPath = Path.GetDirectoryName(ofd.FileName) // converts sub-documents (if any)
                    };
                    converter.Convert(ofd.FileName, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void DocxToMarkdown_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Word OpenXML document|*.docx;*.dotx;*.docm;*.dotm",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Markdown|*.md;*.markdown;*.mkd;*.mkdn;*.mkdwn;*.mdwn;*.mdown;*.markdn;*.mdtxt;*.mdtext",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".md"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var converter = new DocxToMarkdownConverter()
                    {
                        ImagesOutputFolder = Path.GetDirectoryName(sfd.FileName),
                        ImagesBaseUriOverride = "",
                        ImageConverter = new SystemDrawingConverter(), // Converts TIFF, WMF and EMF
                                                                       // (ImageSharp does not support WMF / EMF yet)
                        OriginalFolderPath = Path.GetDirectoryName(ofd.FileName) // converts sub-documents (if any)
                    };
                    converter.Convert(ofd.FileName, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void DocxToMarkdownAppend_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Word OpenXML document|*.docx;*.dotx;*.docm;*.dotm",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Markdown|*.md;*.markdown;*.mkd;*.mkdn;*.mkdwn; *.mdwn;*.mdown;*.markdn;*.mdtxt;*.mdtext",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".md"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var converter = new DocxToMarkdownConverter()
                    {
                        ImagesOutputFolder = Path.GetDirectoryName(sfd.FileName),
                        ImagesBaseUriOverride = "",
                        ImageConverter = new SystemDrawingConverter(), // Converts TIFF, WMF and EMF
                                                                       // (ImageSharp does not support WMF / EMF yet)
                        OriginalFolderPath = Path.GetDirectoryName(ofd.FileName) // converts sub-documents (if any)
                    };
                    converter.Append(ofd.FileName, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void DocxToTxt_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Word OpenXML document|*.docx;*.dotx",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Plain text|*.txt",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".txt"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var converter = new DocxToTxtConverter()
                    {
                        OriginalFolderPath = Path.GetDirectoryName(ofd.FileName) // converts sub-documents (if any)
                    };
                    converter.Convert(ofd.FileName, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void DocxToPdf_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Word OpenXML document|*.docx;*.dotx;*.docm;*.dotm",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Portable Document Format|*.pdf",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".pdf"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var converter = new DocxRenderer()
                    {
                    };
                    converter.SaveAsPdf(ofd.FileName, sfd.FileName);
                }
                catch (Exception ex)
                {                    
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void DocxToImages_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Word OpenXML document|*.docx;*.dotx;*.docm;*.dotm",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new OpenFolderDialog()
            {
                Title = "Choose folder to export images in"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    // var converter = new DocxRenderer()
                    // {                        
                    // };
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void MarkdownToDocx_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Markdown|*.md;*.markdown;*.mkd;*.mkdn;*.mkdwn; *.mdwn;*.mdown;*.markdn;*.mdtxt;*.mdtext",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Word OpenXML document|*.docx;*.dotx;*.docm;*.dotm",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".docx"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var markdown = MarkdownSource.FromFile(ofd.FileName);
                    var converter = new MarkdownConverter()
                    {
                        ImagesBaseUri = Path.GetDirectoryName(ofd.FileName),
                        ImageConverter = new ImageSharpConverter() // Convert WEBP images which are not supported in DOCX
                                                                   // (possibly AVIF and JXL too in a future release) 
                    };
                    converter.ToDocx(markdown, sfd.FileName, FileFormatHelpers.ExtensionToDocumentType(Path.GetExtension(sfd.FileName)));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void MarkdownToDocxAppend_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Markdown|*.md;*.markdown;*.mkd;*.mkdn;*.mkdwn; *.mdwn;*.mdown;*.markdn;*.mdtxt;*.mdtext",
            Title = "Choose the Markdown document to convert",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var ofd2 = new OpenFileDialog()
            {
                Filter = "Word OpenXML document|*.docx;*.dotx;*.docm;*.dotm",
                Title = "Choose the target Word document"
            };
            if (ofd2.ShowDialog(this) == true)
            {
                try
                {
                    var markdown = MarkdownSource.FromFile(ofd.FileName);
                    var converter = new MarkdownConverter()
                    {
                        ImagesBaseUri = Path.GetDirectoryName(ofd.FileName),
                        ImageConverter = new ImageSharpConverter()
                    };
                    converter.ToDocx(markdown, ofd2.FileName, append: true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void MarkdownToRtf_Click(object sender, RoutedEventArgs e)
    {
        // Currently achieved through a two steps conversion.
        var ofd = new OpenFileDialog()
        {
            Filter = "Markdown|*.md;*.markdown;*.mkd;*.mkdn;*.mkdwn; *.mdwn;*.mdown;*.markdn;*.mdtxt;*.mdtext",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Rich Text Format|*.rtf",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".rtf"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var markdown = MarkdownSource.FromFile(ofd.FileName);
                    var converter = new MarkdownConverter()
                    {
                        LinksBaseUri = Path.GetDirectoryName(ofd.FileName), // this will make links absolute
                        ImagesBaseUri = Path.GetDirectoryName(ofd.FileName),
                        ImageConverter = new ImageSharpConverter() // Convert WEBP and GIF images which are not supported in RTF
                                                                   // (possibly AVIF and JXL too in a future release) 
                    };
                    converter.ToRtf(markdown, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void ViewDocx_Click(object sender, RoutedEventArgs e)
    {
        // Please note that the WPF RichTextBox supports a subset of RTF features.
        // To test the DOCX --> RTF conversion provided by DocSharp,
        // the RTF document should be opened in MS Word.
        var ofd = new OpenFileDialog()
        {
            Filter = "Word OpenXML document|*.docx",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            try
            {
                using (var ms = new MemoryStream())
                {
                    var converter = new DocxToRtfConverter()
                    {
                        ImageConverter = new ImageSharpConverter()
                    };
                    converter.Convert(ofd.FileName, ms);
                    var rtbWindow = new Window()
                    {
                        Owner = this,
                        WindowStartupLocation = WindowStartupLocation.CenterOwner
                    };
                    var rtb = new RichTextBox()
                    {
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        VerticalAlignment = VerticalAlignment.Stretch,
                        IsInactiveSelectionHighlightEnabled = true,
                        AutoWordSelection = false,
                        AcceptsReturn = true,
                        AcceptsTab = true,
                    };
                    rtbWindow.Content = rtb;
                    rtb.SelectAll();
                    rtb.Selection.Load(ms, DataFormats.Rtf);
                    rtbWindow.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }

    private async void RtfToDocx_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Rich Text Format|*.rtf",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Word OpenXML document|*.docx",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".docx"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var conv = new RtfToDocxConverter()
                    {                         
                    };
                    conv.Convert(ofd.FileName, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private async void RtfToHtml_Click(object sender, RoutedEventArgs e)
    {
        // Please note that other libraries exist to convert RTF to HTML directly (e.g. RtfPipe), 
        // but they are mostly not maintained. These can be used for internal tests/comparison; 
        // DocSharp RTF -> DOCX -> HTML should produce similar or better results. 
        // There is no plan of adding direct RTF --> HTML to DocSharp because: 
        // - RTF is a complex format that is more similar to DOCX and DOC than HTML. 
        // It supports most Word features, so mapping to DOCX is more natural, 
        // and converting DOCX to HTML is more reliable thanks to the XML-based enumeration. 
        // Libraries such as RtfPipe, while useful for simple RTF, 
        // end up being inaccurate and hard to troubleshoot for complex RTF 
        // (if an advanced RTF feature is not translated to HTML, understanding if the issues lies in the parser
        // or in the converter is not easy, and often requires reworking large parts of the codebase).
        // - To avoid duplicated work we would end up creating an intermediate DOM 
        // (RTF -> DOM -> DOCX/HTML), losing the performance advantage over two-step conversion anyway. 
        // Since we don't the intermediate DOCX to file during the RTF to HTML conversion, 
        // the Open XML document effectively already works as an in-memory DOM 
        // that we don't have to implement from scratch. 
        // As written before, implementing a full RTF object model would be almost as complex as a 
        // DOCX object model, thus reinventing parts of the Open XML SDK. 
        var ofd = new OpenFileDialog()
        {
            Filter = "Rich Text Format|*.rtf",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "HTML|*.html;*.htm",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".html"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var conv = new RtfToDocxConverter();
                    using (var ms = new MemoryStream())
                    {
                        var wpd = conv.ConvertToWordProcessingDocument(ofd.FileName, ms);
                        wpd.SaveTo(sfd.FileName, new HtmlSaveOptions()
                        {
                            ImageConverter = new SystemDrawingConverter(),
                        });
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private async void RtfToMarkdown_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Rich Text Format|*.rtf",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Markdown|*.md;*.markdown;*.mkd;*.mkdn;*.mkdwn;*.mdwn;*.mdown;*.markdn;*.mdtxt;*.mdtext",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".md"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var conv = new RtfToDocxConverter();
                    using (var ms = new MemoryStream())
                    {
                        var wpd = conv.ConvertToWordProcessingDocument(ofd.FileName, ms);
                        wpd.SaveTo(sfd.FileName, new MarkdownSaveOptions()
                        {
                            ImageConverter = new SystemDrawingConverter(),
                            ImagesOutputFolder = Path.GetDirectoryName(sfd.FileName),
                            ImagesBaseUriOverride = "",
                        });
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void DocToRtf_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Multiselect = true,
            Filter = "Microsoft Word 97-2003 document|*.doc;*.dot",
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Rich Text Format|*.rtf",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".rtf"
            };
            if (sfd.ShowDialog(this) == true)
            {
                var tempFile = Path.GetTempFileName();
                try
                {
                    using (var reader = new StructuredStorageReader(ofd.FileName))
                    {
                        var doc = new WordDocument(reader);
                        using (var docx = DocSharp.Binary.OpenXmlLib.WordprocessingML.WordprocessingDocument.Create(tempFile, DocSharp.Binary.OpenXmlLib.WordprocessingDocumentType.Document))
                        {
                            DocSharp.Binary.WordprocessingMLMapping.Converter.Convert(doc, docx);
                        }
                    }
                    var converter = new DocxToRtfConverter()
                    {
                        ImageConverter = new ImageSharpConverter()
                    };
                    converter.Convert(tempFile, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    File.Delete(tempFile);
                }
            }
        }
    }

    private void DocToHtml_Click(object sender, RoutedEventArgs e)
    {
        // Convert DOC to DOCX and then DOCX to HTML.
        var ofd = new OpenFileDialog()
        {
            Multiselect = true,
            Filter = "Microsoft Word 97-2003 document|*.doc;*.dot",
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "HTML|*.html;*.htm",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".htm"
            };
            if (sfd.ShowDialog(this) == true)
            {
                var tempFile = Path.GetTempFileName();
                try
                {
                    using (var reader = new StructuredStorageReader(ofd.FileName))
                    {
                        var doc = new WordDocument(reader);
                        using (var docx = DocSharp.Binary.OpenXmlLib.WordprocessingML.WordprocessingDocument.Create(tempFile, DocSharp.Binary.OpenXmlLib.WordprocessingDocumentType.Document))
                        {
                            DocSharp.Binary.WordprocessingMLMapping.Converter.Convert(doc, docx);
                        }
                    }
                    var converter = new DocxToHtmlConverter()
                    {
                        ImageConverter = new ImageSharpConverter()
                    };
                    converter.Convert(tempFile, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    File.Delete(tempFile);
                }
            }
        }
    }

    private async void HtmlToRtf_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "HTML|*.html;*.htm",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Rich Text Format|*.rtf",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".rtf"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    string html = File.ReadAllText(ofd.FileName);
                    // Move styles inline for better results.
                    var result = PreMailer.Net.PreMailer.MoveCssInline(html);
                    string normalizedHtml = result.Html;

                    string docxFilePath = Path.ChangeExtension(sfd.FileName, ".docx");
                    using (var package = WordprocessingDocument.Create(docxFilePath, WordprocessingDocumentType.Document))
                    {
                        var mainPart = package.AddMainDocumentPart();
                        mainPart.Document = new W.Document();
                        mainPart.Document.AddChild(new W.Body());
                        var htmlConverter = new HtmlConverter(mainPart);
                        await htmlConverter.ParseBody(normalizedHtml);
                        package.Save();
                    }

                    var converter = new DocxToRtfConverter()
                    {
                        ImageConverter = new ImageSharpConverter()
                    };
                    converter.Convert(docxFilePath, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void XlsToHtml_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Multiselect = true,
            Filter = "Microsoft Excel 97-2003 spreadsheet|*.xls;*.xlt;*.xlr",
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "HTML|*.html;*.htm",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".html"
            };
            if (sfd.ShowDialog(this) == true)
            {
                var tempFile = Path.GetTempFileName();
                try
                {
                    using (var reader = new StructuredStorageReader(ofd.FileName))
                    {
                        var xls = new XlsDocument(reader);
                        using (var xlsx = DocSharp.Binary.OpenXmlLib.SpreadsheetML.SpreadsheetDocument.Create(tempFile, DocSharp.Binary.OpenXmlLib.SpreadsheetDocumentType.Workbook))
                        {
                            DocSharp.Binary.SpreadsheetMLMapping.Converter.Convert(xls, xlsx);
                        }
                    }
                    using (var outputStream = new FileStream(sfd.FileName, FileMode.Create, FileAccess.ReadWrite))
                    {
                        XlsxToHtmlConverter.Converter.ConvertXlsx(tempFile, outputStream);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    File.Delete(tempFile);
                }
            }
        }
    }

    private void XlsxToPdf_Click(object sender, RoutedEventArgs e)
    {
         var ofd = new OpenFileDialog()
        {
            Filter = "Excel OpenXML document|*.xlsx;*.xltx;*.xlsm;*.xltm",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "Portable Document Format|*.pdf",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".pdf"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var converter = new XlsxRenderer()
                    {
                    };
                    converter.SaveAsPdf(ofd.FileName, sfd.FileName);
                }
                catch (Exception ex)
                {                    
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void GenerateDocx_Click(object sender, RoutedEventArgs e)
    {
        var sfd = new SaveFileDialog()
        {
            Filter = "Word OpenXML document|*.docx",
        };
        if (sfd.ShowDialog(this) == true)
        {
            try
            {
                string rtfFilePath = Path.ChangeExtension(sfd.FileName, ".rtf");
                string htmlFilePath = Path.ChangeExtension(sfd.FileName, ".html");
                string pdfFilePath = Path.ChangeExtension(sfd.FileName, ".pdf");

                using (var wordDocument = WordprocessingDocument.Create(sfd.FileName, WordprocessingDocumentType.Document))
                {
                    // Add main document part and body
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new W.Document();
                    var body = new W.Body();
                    mainPart.Document.Append(body);

                    // Add a paragraph with formatted text
                    var par = new W.Paragraph();

                    // Add bold text
                    var runBold = new W.Run();
                    runBold.Append(new W.RunProperties(new W.Bold()));
                    runBold.Append(new W.Text("This is bold text. "));
                    par.Append(runBold);

                    // Add italic text
                    var runItalic = new W.Run();
                    runItalic.Append(new W.RunProperties(new W.Italic()));
                    runItalic.Append(new W.Text("This is italic text. "));
                    par.Append(runItalic);

                    // Add text with custom font color
                    var runColored = new W.Run();
                    runColored.Append(new W.RunProperties(new W.Color() { Val = "FF5733" })); // Red
                    runColored.Append(new W.Text("This text has a custom color. "));
                    par.Append(runColored);

                    // Set paragraph alignment
                    var justification = new W.Justification() { Val = W.JustificationValues.Center };
                    var paraProperties = new W.ParagraphProperties();
                    paraProperties.Append(justification);
                    par.Append(paraProperties);
                    
                    // Add paragraph to document body
                    body.Append(par);
                    
                    // Save DOCX document (in the location specified in WordprocessingDocument.Create)
                    mainPart.Document.Save();
                    
                    // Convert document to RTF
                    wordDocument.SaveTo(rtfFilePath, new RtfSaveOptions()
                    {
                        ImageConverter = new ImageSharpConverter(), // Converts TIFF, GIF and other formats which are not supported in RTF.
                    });
                    // Convert document to HTML
                    wordDocument.SaveTo(htmlFilePath, new HtmlSaveOptions()
                    {
                        ExportHeaderFooter = true,
                        ExportFootnotesEndnotes = true,
                        ImageConverter = new SystemDrawingConverter(), // Converts TIFF, WMF and EMF
                                                                       // (ImageSharp does not support WMF / EMF yet)
                    });
                    // Render document to PDF
                    var converter = new DocxRenderer()
                    {
                    };
                    converter.SaveAsPdf(wordDocument, pdfFilePath);
                }
                
            }
            catch (Exception ex)
            {                    
                MessageBox.Show(ex.Message);
            }
        }
    }

    private void GenerateMigraDoc_Click(object sender, RoutedEventArgs e)
    {
        var sfd = new SaveFileDialog()
        {
            Filter = "Rich Text Format|*.rtf",
        };
        if (sfd.ShowDialog(this) == true)
        {
            try
            {
                string pdfFilePath = Path.ChangeExtension(sfd.FileName, ".pdf");
                string docxFilePath = Path.ChangeExtension(sfd.FileName, ".docx");
                
                // Create a MigraDoc document
                var document = new MigraDoc.DocumentObjectModel.Document();
                document.Info.Title = "MigraDoc + DocSharp sample document";
                document.Info.Author = "Author";

                // Add a section
                var section = document.AddSection();
                section.PageSetup.PageFormat = MigraDoc.DocumentObjectModel.PageFormat.A4;
                section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Portrait;

                // Add a paragraph
                var paragraph = section.AddParagraph("Welcome to MigraDoc!");
                paragraph.Format.Font.Bold = true;
                paragraph.Format.Font.Size = 16;

                // Add another paragraph
                var paragraph2 = section.AddParagraph("This is a paragraph with bold and italic text.");
                paragraph2.Format.Font.Italic = true;
                paragraph2.Format.Font.Size = 12;

                // Add a paragraph with custom color
                var paragraph3 = section.AddParagraph("This text has a custom color.");
                paragraph3.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.Red;
                
                // Save as RTF
                var rtfRenderer = new MigraDoc.RtfRendering.RtfDocumentRenderer();
                rtfRenderer.Render(document, sfd.FileName, Path.GetDirectoryName(sfd.FileName) ?? Path.GetTempPath());

                // Save as PDF
                var pdfRenderer = new MigraDoc.Rendering.PdfDocumentRenderer
                {
                    Document = document
                };
                pdfRenderer.RenderDocument();
                pdfRenderer.PdfDocument.Save(pdfFilePath);

                // Convert to DOCX
                var converter = new RtfToDocxConverter();
                converter.Convert(sfd.FileName, docxFilePath);
            }
            catch (Exception ex)
            {                    
                MessageBox.Show(ex.Message);
            }
        }
    }

    private void GenerateXlsx_Click(object sender, RoutedEventArgs e)
    {
        var sfd = new SaveFileDialog()
        {
            Filter = "Excel OpenXML document|*.xlsx",
        };
        if (sfd.ShowDialog(this) == true)
        {
            try
            {
                string pdfFilePath = Path.ChangeExtension(sfd.FileName, ".pdf");
                // Create a new XLSX workbook using ClosedXML
                using (var workbook = new ClosedXML.Excel.XLWorkbook())
                {
                    // Add a new worksheet
                    var worksheet = workbook.Worksheets.Add("Foglio1");

                    // Add values to the first 5 rows and 3 columns of the sheet
                    worksheet.Cell(1, 1).Value = "Header 1";
                    worksheet.Cell(1, 2).Value = "Header 2";
                    worksheet.Cell(1, 3).Value = "Header 3";

                    worksheet.Cell(2, 1).Value = "Value 1A";
                    worksheet.Cell(2, 2).Value = "Value 1B";
                    worksheet.Cell(2, 3).Value = "Value 1C";

                    worksheet.Cell(3, 1).Value = "Value 2A";
                    worksheet.Cell(3, 2).Value = "Value 2B";
                    worksheet.Cell(3, 3).Value = "Value 2C";

                    worksheet.Cell(4, 1).Value = "Value 3A";
                    worksheet.Cell(4, 2).Value = "Value 3B";
                    worksheet.Cell(4, 3).Value = "Value 3C";

                    worksheet.Cell(5, 1).Value = "Value 4A";
                    worksheet.Cell(5, 2).Value = "Value 4B";
                    worksheet.Cell(5, 3).Value = "Value 4C";

                    // Apply formatting to the header row
                    var headerRange = worksheet.Range(1, 1, 1, 3);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGray;
                    headerRange.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;

                    // Save file
                    workbook.SaveAs(sfd.FileName);

                }
                // Render document to PDF
                var converter = new XlsxRenderer()
                {
                };
                converter.SaveAsPdf(sfd.FileName, pdfFilePath);
                
            }
            catch (Exception ex)
            {                    
                MessageBox.Show(ex.Message);
            }
        }
    }
}
