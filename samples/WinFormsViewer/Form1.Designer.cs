namespace WinFormsViewer;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
        toolStrip1 = new ToolStrip();
        newToolStripButton = new ToolStripButton();
        openToolStripButton = new ToolStripButton();
        saveToolStripButton = new ToolStripButton();
        toolStripSeparator = new ToolStripSeparator();
        cutToolStripButton = new ToolStripButton();
        copyToolStripButton = new ToolStripButton();
        pasteToolStripButton = new ToolStripButton();
        toolStripSeparator1 = new ToolStripSeparator();
        boldToolStripButton = new ToolStripButton();
        italicToolStripButton = new ToolStripButton();
        underlineToolStripButton = new ToolStripButton();
        strikeToolStripButton = new ToolStripButton();
        richTextBox1 = new RichTextBox();
        toolStrip1.SuspendLayout();
        SuspendLayout();
        // 
        // toolStrip1
        // 
        toolStrip1.GripStyle = ToolStripGripStyle.Hidden;
        toolStrip1.ImageScalingSize = new Size(20, 20);
        toolStrip1.Items.AddRange(new ToolStripItem[] { newToolStripButton, openToolStripButton, saveToolStripButton, toolStripSeparator, cutToolStripButton, copyToolStripButton, pasteToolStripButton, toolStripSeparator1, boldToolStripButton, italicToolStripButton, underlineToolStripButton, strikeToolStripButton });
        toolStrip1.Location = new Point(0, 0);
        toolStrip1.Name = "toolStrip1";
        toolStrip1.Size = new Size(800, 27);
        toolStrip1.TabIndex = 1;
        toolStrip1.Text = "toolStrip1";
        // 
        // newToolStripButton
        // 
        newToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.Image;
        newToolStripButton.Image = (Image)resources.GetObject("newToolStripButton.Image");
        newToolStripButton.ImageTransparentColor = Color.Magenta;
        newToolStripButton.Name = "newToolStripButton";
        newToolStripButton.Size = new Size(29, 24);
        newToolStripButton.Text = "&New";
        newToolStripButton.Click += newToolStripButton_Click;
        // 
        // openToolStripButton
        // 
        openToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.Image;
        openToolStripButton.Image = (Image)resources.GetObject("openToolStripButton.Image");
        openToolStripButton.ImageTransparentColor = Color.Magenta;
        openToolStripButton.Name = "openToolStripButton";
        openToolStripButton.Size = new Size(29, 24);
        openToolStripButton.Text = "&Open";
        openToolStripButton.ToolTipText = "Open";
        openToolStripButton.Click += openToolStripButton_Click;
        // 
        // saveToolStripButton
        // 
        saveToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.Image;
        saveToolStripButton.Image = (Image)resources.GetObject("saveToolStripButton.Image");
        saveToolStripButton.ImageTransparentColor = Color.Magenta;
        saveToolStripButton.Name = "saveToolStripButton";
        saveToolStripButton.Size = new Size(29, 24);
        saveToolStripButton.Text = "&Save";
        saveToolStripButton.Click += saveToolStripButton_Click;
        // 
        // toolStripSeparator
        // 
        toolStripSeparator.Name = "toolStripSeparator";
        toolStripSeparator.Size = new Size(6, 27);
        // 
        // cutToolStripButton
        // 
        cutToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.Image;
        cutToolStripButton.Image = (Image)resources.GetObject("cutToolStripButton.Image");
        cutToolStripButton.ImageTransparentColor = Color.Magenta;
        cutToolStripButton.Name = "cutToolStripButton";
        cutToolStripButton.Size = new Size(29, 24);
        cutToolStripButton.Text = "&Cut";
        cutToolStripButton.Click += cutToolStripButton_Click;
        // 
        // copyToolStripButton
        // 
        copyToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.Image;
        copyToolStripButton.Image = (Image)resources.GetObject("copyToolStripButton.Image");
        copyToolStripButton.ImageTransparentColor = Color.Magenta;
        copyToolStripButton.Name = "copyToolStripButton";
        copyToolStripButton.Size = new Size(29, 24);
        copyToolStripButton.Text = "&Copy";
        copyToolStripButton.Click += copyToolStripButton_Click;
        // 
        // pasteToolStripButton
        // 
        pasteToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.Image;
        pasteToolStripButton.Image = (Image)resources.GetObject("pasteToolStripButton.Image");
        pasteToolStripButton.ImageTransparentColor = Color.Magenta;
        pasteToolStripButton.Name = "pasteToolStripButton";
        pasteToolStripButton.Size = new Size(29, 24);
        pasteToolStripButton.Text = "&Paste";
        pasteToolStripButton.Click += pasteToolStripButton_Click;
        // 
        // toolStripSeparator1
        // 
        toolStripSeparator1.Name = "toolStripSeparator1";
        toolStripSeparator1.Size = new Size(6, 27);
        // 
        // boldToolStripButton
        // 
        boldToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.Text;
        boldToolStripButton.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        boldToolStripButton.Image = (Image)resources.GetObject("boldToolStripButton.Image");
        boldToolStripButton.ImageTransparentColor = Color.Magenta;
        boldToolStripButton.Name = "boldToolStripButton";
        boldToolStripButton.Size = new Size(29, 24);
        boldToolStripButton.Text = "B";
        boldToolStripButton.ToolTipText = "Toggle bold for selected text";
        boldToolStripButton.Click += boldToolStripButton_Click;
        // 
        // italicToolStripButton
        // 
        italicToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.Text;
        italicToolStripButton.Font = new Font("Segoe UI", 9F, FontStyle.Italic);
        italicToolStripButton.Image = (Image)resources.GetObject("italicToolStripButton.Image");
        italicToolStripButton.ImageTransparentColor = Color.Magenta;
        italicToolStripButton.Name = "italicToolStripButton";
        italicToolStripButton.Size = new Size(29, 24);
        italicToolStripButton.Text = "I";
        italicToolStripButton.ToolTipText = "Toggle italic for selected text";
        italicToolStripButton.Click += italicToolStripButton_Click;
        // 
        // underlineToolStripButton
        // 
        underlineToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.Text;
        underlineToolStripButton.Font = new Font("Segoe UI", 9F, FontStyle.Underline);
        underlineToolStripButton.Image = (Image)resources.GetObject("underlineToolStripButton.Image");
        underlineToolStripButton.ImageTransparentColor = Color.Magenta;
        underlineToolStripButton.Name = "underlineToolStripButton";
        underlineToolStripButton.Size = new Size(29, 24);
        underlineToolStripButton.Text = "U";
        underlineToolStripButton.ToolTipText = "Toggle underline for selected text";
        underlineToolStripButton.Click += underlineToolStripButton_Click;
        // 
        // strikeToolStripButton
        // 
        strikeToolStripButton.DisplayStyle = ToolStripItemDisplayStyle.Text;
        strikeToolStripButton.Font = new Font("Segoe UI", 9F, FontStyle.Strikeout);
        strikeToolStripButton.Image = (Image)resources.GetObject("strikeToolStripButton.Image");
        strikeToolStripButton.ImageTransparentColor = Color.Magenta;
        strikeToolStripButton.Name = "strikeToolStripButton";
        strikeToolStripButton.Size = new Size(29, 24);
        strikeToolStripButton.Text = "S";
        strikeToolStripButton.ToolTipText = "Toggle strikethrough for selected text";
        strikeToolStripButton.Click += strikeToolStripButton_Click;
        // 
        // richTextBox1
        // 
        richTextBox1.BorderStyle = BorderStyle.FixedSingle;
        richTextBox1.Dock = DockStyle.Fill;
        richTextBox1.EnableAutoDragDrop = true;
        richTextBox1.HideSelection = false;
        richTextBox1.Location = new Point(0, 27);
        richTextBox1.Name = "richTextBox1";
        richTextBox1.Size = new Size(800, 423);
        richTextBox1.TabIndex = 2;
        richTextBox1.Text = "";
        richTextBox1.SelectionChanged += richTextBox1_SelectionChanged;
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(8F, 20F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(800, 450);
        Controls.Add(richTextBox1);
        Controls.Add(toolStrip1);
        Name = "Form1";
        Text = "Form1";
        toolStrip1.ResumeLayout(false);
        toolStrip1.PerformLayout();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private ToolStrip toolStrip1;
    private ToolStripButton newToolStripButton;
    private ToolStripButton openToolStripButton;
    private ToolStripButton saveToolStripButton;
    private ToolStripSeparator toolStripSeparator;
    private ToolStripButton cutToolStripButton;
    private ToolStripButton copyToolStripButton;
    private ToolStripButton pasteToolStripButton;
    private ToolStripSeparator toolStripSeparator1;
    private ToolStripButton boldToolStripButton;
    private ToolStripButton italicToolStripButton;
    private ToolStripButton underlineToolStripButton;
    private ToolStripButton strikeToolStripButton;
    private RichTextBox richTextBox1;
}
