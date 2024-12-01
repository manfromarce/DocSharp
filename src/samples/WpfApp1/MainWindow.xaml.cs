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
using DocSharp.Docx;
using Microsoft.Win32;
using b2xtranslator.StructuredStorage.Reader;
using WordprocessingDocument = b2xtranslator.OpenXmlLib.WordprocessingML.WordprocessingDocument;
using SpreadsheetDocument = b2xtranslator.OpenXmlLib.SpreadsheetML.SpreadsheetDocument;
using PresentationDocument = b2xtranslator.OpenXmlLib.PresentationML.PresentationDocument;
using b2xtranslator.DocFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.PptFileFormat;
using DocSharp.Markdown;

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
                string outputDir = folderDlg.FolderName;
                foreach (string file in ofd.FileNames)
                {
                    string inputExt = Path.GetExtension(file).ToLower();
                    try
                    {
                        using (var reader = new StructuredStorageReader(file))
                        {
                            string outputExt = inputExt + "x";
                            string baseName = Path.GetFileNameWithoutExtension(file);
                            string outputFile = Path.Join(outputDir, baseName + outputExt);
                            var outputType = b2xtranslator.OpenXmlLib.OpenXmlDocumentType.Document;
                            if (inputExt == ".dot" || inputExt == ".xlt" || inputExt == ".pot")
                            {
                                outputType = b2xtranslator.OpenXmlLib.OpenXmlDocumentType.Template;
                            }
                            switch (inputExt)
                            {
                                case ".doc":
                                case ".dot":
                                    var doc = new WordDocument(reader);
                                    using (var docx = WordprocessingDocument.Create(outputFile, outputType))
                                    {
                                        b2xtranslator.WordprocessingMLMapping.Converter.Convert(doc, docx);
                                    }
                                    break;
                                case ".xls":
                                case ".xlt":
                                    var xls = new XlsDocument(reader);
                                    using (var xlsx = SpreadsheetDocument.Create(outputFile, outputType))
                                    {
                                        b2xtranslator.SpreadsheetMLMapping.Converter.Convert(xls, xlsx);
                                    }
                                    break;
                                case ".ppt":
                                case ".pps":
                                case ".pot":
                                    var ppt = new PowerpointDocument(reader);
                                    using (var pptx = PresentationDocument.Create(outputFile, outputType))
                                    {
                                        b2xtranslator.PresentationMLMapping.Converter.Convert(ppt, pptx);
                                    }
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw;
                        //MessageBox.Show("Conversion failed: " + Environment.NewLine + ex.Message);
                    }
                }
            }
        }
    }

    private void DocxToMarkdown_Click(object sender, RoutedEventArgs e)
    {
        var ofd = new OpenFileDialog()
        {
            Filter = "Word OpenXML document|*.docx",
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
                var converter = new DocxToMarkdownConverter()
                {
                    ImagesOutputFolder = Path.GetDirectoryName(sfd.FileName),
                    ImagesBaseUriOverride = "",
                    //ImagesBaseUriOverride = "..",
                    //ImagesBaseUriOverride = "../images",
                    //ImagesBaseUriOverride = "../images/",
                    //ImagesBaseUriOverride = @"..\images\",
                    //ImagesBaseUriOverride = "images",
                    //ImagesBaseUriOverride = "images/",
                    //ImagesBaseUriOverride = @"images\",
                    //ImagesBaseUriOverride = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                };
                converter.Convert(ofd.FileName, sfd.FileName);
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
                var markdown = MarkdownSource.FromFile(ofd.FileName);
                MarkdownConverter.ToDocx(markdown, sfd.FileName);
            }
        }
    }
}
