using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;
using Microsoft.Win32;
using WordprocessingDocument = DocSharp.Binary.OpenXmlLib.WordprocessingML.WordprocessingDocument;
using SpreadsheetDocument = DocSharp.Binary.OpenXmlLib.SpreadsheetML.SpreadsheetDocument;
using PresentationDocument = DocSharp.Binary.OpenXmlLib.PresentationML.PresentationDocument;
using WordprocessingDocumentType = DocSharp.Binary.OpenXmlLib.WordprocessingDocumentType;
using SpreadsheetDocumentType = DocSharp.Binary.OpenXmlLib.SpreadsheetDocumentType;
using PresentationDocumentType = DocSharp.Binary.OpenXmlLib.PresentationDocumentType;
using DocSharp.Binary.DocFileFormat;
using DocSharp.Binary.Spreadsheet.XlsFileFormat;
using DocSharp.Binary.PptFileFormat;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Docx;
using DocSharp.Imaging;
using DocSharp.Markdown;
using DocSharp.Renderer;
using DocSharp.Rtf;
using HtmlToOpenXml;
using PeachPDF;
using PeachPDF.Network;
using System.Net.Http;

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
                                    var docxType = inputExt == ".dot" ? WordprocessingDocumentType.Template :
                                                                          WordprocessingDocumentType.Document;
                                    using (var docx = WordprocessingDocument.Create(outputFile, docxType))
                                    {
                                        DocSharp.Binary.WordprocessingMLMapping.Converter.Convert(doc, docx);
                                    }
                                    break;
                                case ".xls":
                                case ".xlt":
                                    var xls = new XlsDocument(reader);
                                    var xlsxType = inputExt == ".xlt" ? SpreadsheetDocumentType.Template :
                                                                         SpreadsheetDocumentType.Workbook;
                                    using (var xlsx = SpreadsheetDocument.Create(outputFile, xlsxType))
                                    {
                                        DocSharp.Binary.SpreadsheetMLMapping.Converter.Convert(xls, xlsx);
                                    }
                                    break;
                                case ".ppt":
                                case ".pps":
                                case ".pot":
                                    var ppt = new PowerpointDocument(reader);
                                    var pptxType = PresentationDocumentType.Presentation;
                                    if (inputExt == ".pot")
                                    {
                                        pptxType = PresentationDocumentType.Template;
                                    }
                                    else if (inputExt == ".pps")
                                    {
                                        pptxType = PresentationDocumentType.Slideshow;
                                    }
                                    using (var pptx = PresentationDocument.Create(outputFile, pptxType))
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
            Filter = "Word OpenXML document|*.docx;*.dotx",
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
                        ImageConverter = new ImageSharpConverter()
                        // Converts TIFF, GIF and other formats which are not supported in RTF.
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
            Filter = "Word OpenXML document|*.docx;*.dotx",
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
                        ImageConverter = new SystemDrawingConverter() // Converts TIFF, WMF and EMF
                                                                      // (ImageSharp does not support WMF / EMF yet)
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
            Filter = "Word OpenXML document|*.docx;*.dotx",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "PDF|*.pdf",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".pdf"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var renderer = new WordRenderer();
                    renderer.ConvertToPdf(ofd.FileName, sfd.FileName);
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
            Filter = "Word OpenXML document|*.docx;*.dotx",
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
                        ExportHeaderFooter = true,
                        ExportFootnotesEndnotes = true,
                        ImagesOutputFolder = Path.GetDirectoryName(sfd.FileName),
                        ImagesBaseUriOverride = "",
                        //ImagesBaseUriOverride = "..",
                        //ImagesBaseUriOverride = "images/",
                        ImageConverter = new SystemDrawingConverter() // Converts TIFF, WMF and EMF
                                                                      // (ImageSharp does not support WMF / EMF yet)
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
                        ExportHeaderFooter = true,
                        ExportFootnotesEndnotes = true
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

    private void RtfToDocx_Click(object sender, RoutedEventArgs e)
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
                    var rtf = RtfSource.FromFile(ofd.FileName);
                    rtf.ToDocx(sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
    
    private void RtfToHtml_Click(object sender, RoutedEventArgs e)
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
                Filter = "HTML|*.html;*.htm",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".html"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var rtf = RtfSource.FromFile(ofd.FileName);
                    rtf.ToHtml(sfd.FileName, new RtfToHtmlSettings()
                    {
                        ImageConverter = new ImageSharpConverter()
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void RtfToMarkdown_Click(object sender, RoutedEventArgs e)
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
                Filter = "Markdown|*.md;*.markdown;*.mkd;*.mkdn;*.mkdwn; *.mdwn;*.mdown;*.markdn;*.mdtxt;*.mdtext",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".md"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var rtf = RtfSource.FromFile(ofd.FileName);
                    rtf.ToMarkdown(sfd.FileName, new RtfToMdSettings()
                    {
                        ImagesOutputFolder = Path.GetDirectoryName(sfd.FileName),
                        ImagesBaseUriOverride = "",
                        ImageConverter = new ImageSharpConverter()
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
                    }
                }

    private void RtfToTxt_Click(object sender, RoutedEventArgs e)
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
                Filter = "Plain text|*.txt",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".txt"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var rtf = RtfSource.FromFile(ofd.FileName);
                    rtf.ToPlainText(sfd.FileName);
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
                Filter = "Word OpenXML document|*.docx",
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
                    converter.ToDocx(markdown, sfd.FileName);
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
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var ofd2 = new SaveFileDialog()
            {
                Filter = "Word OpenXML document|*.docx",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".docx"
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
                        VerticalAlignment = System.Windows.VerticalAlignment.Stretch,
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

    private async void HtmlToRtf_Click(object sender, RoutedEventArgs e)
    {
        // Convert HTML to DOCX using the HtmlToOpenXml library and then DOCX to RTF using DocSharp.
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
                    using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(docxFilePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                    {
                        var mainPart = package.AddMainDocumentPart();
                        mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                        mainPart.Document.AddChild(new DocumentFormat.OpenXml.Wordprocessing.Body());                        
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

    private void DocToRtf_Click(object sender, RoutedEventArgs e)
    {
        // Convert DOC to DOCX and then DOCX to RTF.
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
                        using (var docx = WordprocessingDocument.Create(tempFile, WordprocessingDocumentType.Document))
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
                        using (var docx = WordprocessingDocument.Create(tempFile, WordprocessingDocumentType.Document))
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

    private void XlsToHtml_Click(object sender, RoutedEventArgs e)
    {
        // Convert XLS to XLSX using DocSharp and then XLSX to HTML using the XlsxToHtmlConverter library.
        var ofd = new OpenFileDialog()
        {
            Multiselect = true,
            Filter = "Spreadsheet|*.xlsx;*.xltx;*.xls;*.xlt;*.xlr",
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
                    bool isXls = false;
                    switch (Path.GetExtension(Path.GetFileNameWithoutExtension(ofd.FileName)).ToLower())
                    {
                        case ".xls":
                        case ".xlt":
                        case ".xlr":
                            isXls = true;
                            // Convert XLS to XLSX using DocSharp.Binary if necessary
                    using (var reader = new StructuredStorageReader(ofd.FileName))
                    {
                        var xls = new XlsDocument(reader);
                        using (var xlsx = SpreadsheetDocument.Create(tempFile, SpreadsheetDocumentType.Workbook))
                        {
                            DocSharp.Binary.SpreadsheetMLMapping.Converter.Convert(xls, xlsx);
                        }
                    }
                        break;  
                    }
                    using (var outputStream = new FileStream(sfd.FileName, FileMode.Create, FileAccess.ReadWrite))
                    {
                        XlsxToHtmlConverter.Converter.ConvertXlsx(isXls ? tempFile : ofd.FileName, outputStream);
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

    private async void RtfToPdf_Click(object sender, RoutedEventArgs e)
    {
        // RTF to PDF is not directly supported, but there are two ways to achieve it: 
        // - Convert RTF to DOCX and then DOCX to PDF 
        // - Convert RTF to HTML using DocSharp and then HTML to PDF using the PeachPdf library
        // Since the RTF to DOCX conversion is still experimental, the second way is recommended at this time.
        var ofd = new OpenFileDialog()
        {
            Filter = "Rich Text Format|*.rtf",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "PDF|*.pdf",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".pdf"
            };
            if (sfd.ShowDialog(this) == true)
            {
                try
                {
                    var rtf = RtfSource.FromFile(ofd.FileName);
                    string html = rtf.ToHtml(new RtfToHtmlSettings()
                    {
                        ImageConverter = new ImageSharpConverter()
                    });
                    var pdfConfig = new PdfGenerateConfig()
                    {
                        PageSize = PeachPDF.PdfSharpCore.PageSize.Letter,
                        PageOrientation = PeachPDF.PdfSharpCore.PageOrientation.Portrait
                    };
                    var generator = new PdfGenerator();
                    using (var document = await generator.GeneratePdf(html, pdfConfig))
                    {
                        document.Save(sfd.FileName);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private async void HtmlToPdf_Click(object sender, RoutedEventArgs e)
    {
        // For HTML to PDF you can use the PeachPDF library.
        // DocSharp also uses the PeachPDF.PdfSharpCore package for DOCX to PDF conversion,
        // so that few dependencies are needed.
        var ofd = new OpenFileDialog()
        {
            Filter = "HTML|*.html;*.htm;*.mhtml;*.mht",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var sfd = new SaveFileDialog()
            {
                Filter = "PDF|*.pdf",
                FileName = Path.GetFileNameWithoutExtension(ofd.FileName) + ".pdf"
            };
            if (sfd.ShowDialog(this) == true)
            {
                bool isMhtml = Path.GetExtension(ofd.FileName).ToLower() switch
                {
                    ".mhtml" => true,
                    ".mht" => true,
                    _ => false
                };
                FileStream? mhtmlStream = null;
                var currentDir = Directory.GetCurrentDirectory();
                try
                {
                    // For MHTML archives, pass null to GeneratePdf to load the content from the NetworkLoader instead.
                    var html = isMhtml ? null : File.ReadAllText(ofd.FileName);
                    if (isMhtml)
                    {
                        // Open stream to the MHTML file.
                        mhtmlStream = File.OpenRead(ofd.FileName);
                    }
                    using (var httpClient = new HttpClient())
                    {
                        // Fixes issue with servers refusing connections from clients without a user agent
                        httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36 Edg/137.0.0.0");
                        httpClient.Timeout = TimeSpan.FromSeconds(30);

                        // Set directory to process local files (images, stylesheets, ...) referenced in the HTML page.
                        var directory = Path.GetFullPath(Path.GetDirectoryName(ofd.FileName) ?? currentDir);
                        Directory.SetCurrentDirectory(directory);

                        var pdfConfig = new PdfGenerateConfig()
                        {
                            PageSize = PeachPDF.PdfSharpCore.PageSize.Letter,
                            PageOrientation = PeachPDF.PdfSharpCore.PageOrientation.Portrait,
                            // For regular HTML, allow access to online resources;
                            // for MHTML, a special NetworkLoader is needed.
                            NetworkLoader = isMhtml ? new MimeKitNetworkLoader(mhtmlStream) :
                                                      new HttpClientNetworkLoader(httpClient, new Uri(directory))
                        };

                        var generator = new PdfGenerator();
                        using (var document = await generator.GeneratePdf(html, pdfConfig))
                        {
                            document.Save(sfd.FileName);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                // Close the MHTML stream, if any
                mhtmlStream?.Dispose();

                // Restore the current directory
                Directory.SetCurrentDirectory(currentDir);
            }
        }
    }
}
