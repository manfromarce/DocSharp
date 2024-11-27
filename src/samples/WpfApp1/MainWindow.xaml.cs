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

    private void DocxToMdTest_Click(object sender, RoutedEventArgs e)
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
                var converter = new DocxToMarkdownConverter();
                converter.Convert(ofd.FileName, sfd.FileName);
            }
        }
    }
}