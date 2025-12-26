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
using DocSharp.Markdown;
using DocSharp.Imaging;
using HtmlToOpenXml;

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

    private async void RtfToHtmlDocx_Click(object sender, RoutedEventArgs e)
    {
        // RTF to DOCX/HTML is not implemented yet in DocSharp but it's planned.
        // This is a workaround based on other open source libraries and will be used as comparison.
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
                    string html = RtfPipe.Rtf.ToHtml(File.ReadAllText(ofd.FileName));
                    File.WriteAllText(Path.ChangeExtension(sfd.FileName, ".html"), html);
                    using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(sfd.FileName, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                    {
                        var mainPart = package.AddMainDocumentPart();
                        mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                        mainPart.Document.AddChild(new DocumentFormat.OpenXml.Wordprocessing.Body());
                        var htmlConverter = new HtmlConverter(mainPart);
                        await htmlConverter.ParseBody(html);
                        package.Save();
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
                        using (var xlsx = SpreadsheetDocument.Create(tempFile, SpreadsheetDocumentType.Workbook))
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

}
