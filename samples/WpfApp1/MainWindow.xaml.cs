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
using DocSharp.Binary.DocFileFormat;
using DocSharp.Binary.Spreadsheet.XlsFileFormat;
using DocSharp.Binary.PptFileFormat;
using DocSharp.Binary.StructuredStorage.Reader;
using DocSharp.Docx;
using DocSharp.Markdown;
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
            Filter = "Office 97-2003 documents|*.doc;*.dot;*.xls;*.xlt;*.ppt;*.pps;*.pot",
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
                            var outputType = DocSharp.Binary.OpenXmlLib.OpenXmlDocumentType.Document;
                            if (inputExt == ".dot" || inputExt == ".xlt" || inputExt == ".pot")
                            {
                                outputType = DocSharp.Binary.OpenXmlLib.OpenXmlDocumentType.Template;
                            }
                            switch (inputExt)
                            {
                                case ".doc":
                                case ".dot":
                                    var doc = new WordDocument(reader);
                                    using (var docx = WordprocessingDocument.Create(outputFile, outputType))
                                    {
                                        DocSharp.Binary.WordprocessingMLMapping.Converter.Convert(doc, docx);
                                    }
                                    break;
                                case ".xls":
                                case ".xlt":
                                    var xls = new XlsDocument(reader);
                                    using (var xlsx = SpreadsheetDocument.Create(outputFile, outputType))
                                    {
                                        DocSharp.Binary.SpreadsheetMLMapping.Converter.Convert(xls, xlsx);
                                    }
                                    break;
                                case ".ppt":
                                case ".pps":
                                case ".pot":
                                    var ppt = new PowerpointDocument(reader);
                                    using (var pptx = PresentationDocument.Create(outputFile, outputType))
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
                Filter = "Markdown|*.md;*.markdown",
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
                        //ImagesBaseUriOverride = "..",
                        //ImagesBaseUriOverride = "images/",
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
                    var converter = new DocxToRtfConverter();
                    converter.Convert(ofd.FileName, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private async void RtfToDocx_Click(object sender, RoutedEventArgs e)
    {
        // The RTF to DOCX is not implemented yet in DocSharp but it's planned.
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

    private void MarkdownToDocx_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Markdown|*.md;*.markdown",
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
                        ImagesBaseUri = Path.GetDirectoryName(ofd.FileName)
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
            Filter = "Markdown|*.md;*.markdown",
            Multiselect = false,
        };
        if (ofd.ShowDialog(this) == true)
        {
            var ofd2 = new OpenFileDialog()
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
                        ImagesBaseUri = Path.GetDirectoryName(ofd.FileName)
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

                    var converter = new DocxToRtfConverter();
                    converter.Convert(docxFilePath, sfd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }

    private void DocxRtfToHtml_Click(object sender, RoutedEventArgs e)
    {
        // Please note that other open source libraries exist to convert DOCX to HTML directly, 
        // e.g. OpenXmlToHtml (based on a fork of OpenXmlPowerTools) would give better results.
        // This sample is mainly to test the DOCX to RTF conversion
        // and if the produced RTF is correctly interpreted by third-party tools.
        var ofd = new OpenFileDialog()
        {
            Filter = "Documents|*.docx;*.dotx;*.rtf",
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
                    string rtfContent = "";
                    switch (Path.GetExtension(ofd.FileName).ToLower())
                    {
                        case ".docx":
                        case ".dotx":
                            var converter = new DocxToRtfConverter();
                            rtfContent = converter.ConvertToString(ofd.FileName);
                            break;
                        case ".rtf":
                            rtfContent = File.ReadAllText(ofd.FileName);
                            break;
                    }
                    string html = RtfPipe.Rtf.ToHtml(rtfContent);
                    File.WriteAllText(sfd.FileName, html);
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
                    var converter = new DocxToRtfConverter();
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
}
