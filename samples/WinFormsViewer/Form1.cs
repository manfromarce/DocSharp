using DocSharp.Docx;
using DocSharp.Imaging;
using DocSharp.Markdown;

namespace WinFormsViewer;

public partial class Form1 : Form
{
    public Form1()
    {
        InitializeComponent();
    }

    private void newToolStripButton_Click(object sender, EventArgs e)
    {
        var form = new Form1();
        form.Show();
    }

    private void openToolStripButton_Click(object sender, EventArgs e)
    {
        using (var ofd = new OpenFileDialog()
        {
            Filter = "Supported documents|*.rtf;*.docx;*.md;*.txt|Rich Text Format|*.rtf|Microsoft Word document|*.docx;*.docm;*.dotx;*.dotm|Markdown document|*.md;*.markdown|All plain text files|*.*"
        })
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                switch (Path.GetExtension(ofd.FileName).ToLowerInvariant())
                {
                    case ".rtf":
                        richTextBox1.LoadFile(ofd.FileName, RichTextBoxStreamType.RichText);
                        break;
                    case ".docx":
                    case ".docm":
                    case ".dotx":
                    case ".dotm":
                        try
                        {
                            using (var ms = new MemoryStream())
                            {
                                var converter = new DocxToRtfConverter()
                                {
                                    ImageConverter = new SystemDrawingConverter(),
                                };
                                converter.Convert(ofd.FileName, ms);
                                ms.Position = 0;
                                richTextBox1.LoadFile(ms, RichTextBoxStreamType.RichText);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        break;
                    case ".md":
                    case ".markdown":
                        try
                        {
                            using (var ms = new MemoryStream())
                            {
                                var markdown = MarkdownSource.FromFile(ofd.FileName);
                                var converter = new MarkdownConverter()
                                {
                                    ImagesBaseUri = Path.GetDirectoryName(ofd.FileName),
                                    ImageConverter = new SystemDrawingConverter()
                                };
                                converter.ToRtf(markdown, ms);
                                ms.Position = 0;
                                richTextBox1.LoadFile(ms, RichTextBoxStreamType.RichText);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        break;
                    case ".txt":
                    default:
                        richTextBox1.LoadFile(ofd.FileName, RichTextBoxStreamType.PlainText);
                        break;
                }
            }
        }
    }

    private void saveToolStripButton_Click(object sender, EventArgs e)
    {
        using (var sfd = new SaveFileDialog()
        {
            Filter = "Rich Text Format|*.rtf|Plain text|*.txt"
        })
        {
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                switch (Path.GetExtension(sfd.FileName).ToLowerInvariant())
                {
                    case ".rtf":
                        richTextBox1.SaveFile(sfd.FileName, RichTextBoxStreamType.RichText);
                        break;
                    case ".docx":
                    case ".docm":
                    case ".dotx":
                    case ".dotm":
                        break;
                    case ".md":
                    case ".markdown":
                        break;
                    case ".txt":
                    default:
                        richTextBox1.SaveFile(sfd.FileName, RichTextBoxStreamType.UnicodePlainText);
                        break;
                }
            }
        }
    }

    private void cutToolStripButton_Click(object sender, EventArgs e)
    {
        richTextBox1.Cut();
    }

    private void copyToolStripButton_Click(object sender, EventArgs e)
    {
        richTextBox1.Copy();
    }

    private void pasteToolStripButton_Click(object sender, EventArgs e)
    {
        richTextBox1.Paste();
    }

    private void boldToolStripButton_Click(object sender, EventArgs e)
    {
        if (richTextBox1.SelectionFont != null)
        {
            var style = richTextBox1.SelectionFont.Style;
            richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, style ^ FontStyle.Bold);
            UpdateButtons();
        }
    }

    private void italicToolStripButton_Click(object sender, EventArgs e)
    {
        if (richTextBox1.SelectionFont != null)
        {
            var style = richTextBox1.SelectionFont.Style;
            richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, style ^ FontStyle.Italic);
            UpdateButtons();
        }
    }

    private void underlineToolStripButton_Click(object sender, EventArgs e)
    {
        if (richTextBox1.SelectionFont != null)
        {
            var style = richTextBox1.SelectionFont.Style;
            richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, style ^ FontStyle.Underline);
            UpdateButtons();
        }
    }

    private void strikeToolStripButton_Click(object sender, EventArgs e)
    {
        if (richTextBox1.SelectionFont != null)
        {
            var style = richTextBox1.SelectionFont.Style;
            richTextBox1.SelectionFont = new Font(richTextBox1.SelectionFont, style ^ FontStyle.Strikeout);
            UpdateButtons();
        }
    }

    private void UpdateButtons()
    {
        boldToolStripButton.Checked = richTextBox1.SelectionFont?.Bold == true;
        italicToolStripButton.Checked = richTextBox1.SelectionFont?.Italic == true;
        underlineToolStripButton.Checked = richTextBox1.SelectionFont?.Underline == true;
        strikeToolStripButton.Checked = richTextBox1.SelectionFont?.Strikeout == true;
    }

    private void richTextBox1_SelectionChanged(object sender, EventArgs e)
    {
        UpdateButtons();
    }
}
