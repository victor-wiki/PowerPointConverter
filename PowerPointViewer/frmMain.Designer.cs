namespace PowerPointViewer
{
    partial class frmMain
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
            webView = new Microsoft.Web.WebView2.WinForms.WebView2();
            menuStrip1 = new MenuStrip();
            tsmiOpenFile = new ToolStripMenuItem();
            toolsToolStripMenuItem = new ToolStripMenuItem();
            tsmiColorTransform = new ToolStripMenuItem();
            openFileDialog1 = new OpenFileDialog();
            btnFirst = new Button();
            btnPrevious = new Button();
            btnNext = new Button();
            btnLast = new Button();
            cboNumber = new ComboBox();
            label1 = new Label();
            lblTotal = new Label();
            lblMessage = new Label();
            chkEnableLog = new CheckBox();
            chkUseLowQualityForImage = new CheckBox();
            ((System.ComponentModel.ISupportInitialize)webView).BeginInit();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // webView
            // 
            webView.AllowExternalDrop = true;
            webView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            webView.CreationProperties = null;
            webView.DefaultBackgroundColor = Color.White;
            webView.Location = new Point(1, 57);
            webView.Name = "webView";
            webView.Size = new Size(796, 392);
            webView.TabIndex = 0;
            webView.ZoomFactor = 1D;
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { tsmiOpenFile, toolsToolStripMenuItem });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(800, 25);
            menuStrip1.TabIndex = 1;
            menuStrip1.Text = "menuStrip1";
            // 
            // tsmiOpenFile
            // 
            tsmiOpenFile.Name = "tsmiOpenFile";
            tsmiOpenFile.Size = new Size(75, 21);
            tsmiOpenFile.Text = "Open File";
            tsmiOpenFile.Click += tsmiOpenFile_Click;
            // 
            // toolsToolStripMenuItem
            // 
            toolsToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { tsmiColorTransform });
            toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            toolsToolStripMenuItem.Size = new Size(52, 21);
            toolsToolStripMenuItem.Text = "Tools";
            // 
            // tsmiColorTransform
            // 
            tsmiColorTransform.Name = "tsmiColorTransform";
            tsmiColorTransform.Size = new Size(172, 22);
            tsmiColorTransform.Text = "Color Transform";
            tsmiColorTransform.Click += tsmiColorTransform_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            openFileDialog1.Filter = "PPT file|*.pptx";
            // 
            // btnFirst
            // 
            btnFirst.Enabled = false;
            btnFirst.Location = new Point(12, 29);
            btnFirst.Name = "btnFirst";
            btnFirst.Size = new Size(41, 23);
            btnFirst.TabIndex = 2;
            btnFirst.Text = "|<";
            btnFirst.UseVisualStyleBackColor = true;
            btnFirst.Click += btnFirst_Click;
            // 
            // btnPrevious
            // 
            btnPrevious.Enabled = false;
            btnPrevious.Location = new Point(59, 29);
            btnPrevious.Name = "btnPrevious";
            btnPrevious.Size = new Size(41, 23);
            btnPrevious.TabIndex = 3;
            btnPrevious.Text = "<";
            btnPrevious.UseVisualStyleBackColor = true;
            btnPrevious.Click += btnPrevious_Click;
            // 
            // btnNext
            // 
            btnNext.Enabled = false;
            btnNext.Location = new Point(106, 29);
            btnNext.Name = "btnNext";
            btnNext.Size = new Size(41, 23);
            btnNext.TabIndex = 4;
            btnNext.Text = ">";
            btnNext.UseVisualStyleBackColor = true;
            btnNext.Click += btnNext_Click;
            // 
            // btnLast
            // 
            btnLast.Enabled = false;
            btnLast.Location = new Point(153, 29);
            btnLast.Name = "btnLast";
            btnLast.Size = new Size(41, 23);
            btnLast.TabIndex = 5;
            btnLast.Text = ">|";
            btnLast.UseVisualStyleBackColor = true;
            btnLast.Click += btnLast_Click;
            // 
            // cboNumber
            // 
            cboNumber.DropDownStyle = ComboBoxStyle.DropDownList;
            cboNumber.FormattingEnabled = true;
            cboNumber.Location = new Point(202, 29);
            cboNumber.Name = "cboNumber";
            cboNumber.Size = new Size(38, 25);
            cboNumber.TabIndex = 6;
            cboNumber.SelectedIndexChanged += cboNumber_SelectedIndexChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(247, 32);
            label1.Name = "label1";
            label1.Size = new Size(13, 17);
            label1.TabIndex = 7;
            label1.Text = "/";
            // 
            // lblTotal
            // 
            lblTotal.AutoSize = true;
            lblTotal.Location = new Point(265, 32);
            lblTotal.Name = "lblTotal";
            lblTotal.Size = new Size(15, 17);
            lblTotal.TabIndex = 8;
            lblTotal.Text = "0";
            // 
            // lblMessage
            // 
            lblMessage.AutoSize = true;
            lblMessage.Location = new Point(445, 35);
            lblMessage.Name = "lblMessage";
            lblMessage.Size = new Size(0, 17);
            lblMessage.TabIndex = 9;
            // 
            // chkEnableLog
            // 
            chkEnableLog.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            chkEnableLog.AutoSize = true;
            chkEnableLog.Checked = true;
            chkEnableLog.CheckState = CheckState.Checked;
            chkEnableLog.Location = new Point(507, 32);
            chkEnableLog.Name = "chkEnableLog";
            chkEnableLog.Size = new Size(89, 21);
            chkEnableLog.TabIndex = 10;
            chkEnableLog.Text = "Enable log";
            chkEnableLog.UseVisualStyleBackColor = true;
            // 
            // chkUseLowQualityForImage
            // 
            chkUseLowQualityForImage.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            chkUseLowQualityForImage.AutoSize = true;
            chkUseLowQualityForImage.Location = new Point(635, 33);
            chkUseLowQualityForImage.Name = "chkUseLowQualityForImage";
            chkUseLowQualityForImage.Size = new Size(153, 21);
            chkUseLowQualityForImage.TabIndex = 11;
            chkUseLowQualityForImage.Text = "Low quality for image";
            chkUseLowQualityForImage.UseVisualStyleBackColor = true;
            // 
            // frmMain
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(chkUseLowQualityForImage);
            Controls.Add(chkEnableLog);
            Controls.Add(lblMessage);
            Controls.Add(lblTotal);
            Controls.Add(label1);
            Controls.Add(cboNumber);
            Controls.Add(btnLast);
            Controls.Add(btnNext);
            Controls.Add(btnPrevious);
            Controls.Add(btnFirst);
            Controls.Add(webView);
            Controls.Add(menuStrip1);
            MainMenuStrip = menuStrip1;
            Name = "frmMain";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "PPT Viewer";
            WindowState = FormWindowState.Maximized;
            Load += frmMain_Load;
            ((System.ComponentModel.ISupportInitialize)webView).EndInit();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Microsoft.Web.WebView2.WinForms.WebView2 webView;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem tsmiOpenFile;
        private OpenFileDialog openFileDialog1;
        private Button btnFirst;
        private Button btnPrevious;
        private Button btnNext;
        private Button btnLast;
        private ComboBox cboNumber;
        private Label label1;
        private Label lblTotal;
        private Label label2;
        private Label lblMessage;
        private ToolStripMenuItem toolsToolStripMenuItem;
        private ToolStripMenuItem tsmiColorTransform;
        private CheckBox chkEnableLog;
        private CheckBox chkUseLowQualityForImage;
    }
}
