namespace ExcelCustomAddin
{
    partial class ActionPanelControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ActionPanelControl));
            this.metroPanel1 = new MetroFramework.Controls.MetroPanel();
            this.btnResetEvidenceNo = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.cbScalePercent = new MetroFramework.Controls.MetroRadioButton();
            this.cbAutoFixWidth = new MetroFramework.Controls.MetroRadioButton();
            this.numScalePercent = new System.Windows.Forms.NumericUpDown();
            this.btnFormatImages = new System.Windows.Forms.Button();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripFilePath = new System.Windows.Forms.ToolStripLabel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkInsertOnNewPage = new MetroFramework.Controls.MetroCheckBox();
            this.btnInsertPictures = new System.Windows.Forms.Button();
            this.txtImagePath = new MetroFramework.Controls.MetroTextBox();
            this.btnCreateEvidence = new System.Windows.Forms.Button();
            this.btnFormatDocument = new System.Windows.Forms.Button();
            this.listofSheet = new System.Windows.Forms.ListView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.btnChangeSheetName = new System.Windows.Forms.ToolStripMenuItem();
            this.btnPinSheet = new System.Windows.Forms.ToolStripMenuItem();
            this.btnInsertMultipleImages = new System.Windows.Forms.Button();
            this.metroPanel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numScalePercent)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // metroPanel1
            // 
            this.metroPanel1.Controls.Add(this.btnResetEvidenceNo);
            this.metroPanel1.Controls.Add(this.groupBox2);
            this.metroPanel1.Controls.Add(this.btnFormatImages);
            this.metroPanel1.Controls.Add(this.toolStrip1);
            this.metroPanel1.Controls.Add(this.groupBox1);
            this.metroPanel1.Controls.Add(this.btnCreateEvidence);
            this.metroPanel1.Controls.Add(this.btnFormatDocument);
            this.metroPanel1.Controls.Add(this.listofSheet);
            this.metroPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.metroPanel1.HorizontalScrollbarBarColor = true;
            this.metroPanel1.HorizontalScrollbarHighlightOnWheel = false;
            this.metroPanel1.HorizontalScrollbarSize = 10;
            this.metroPanel1.Location = new System.Drawing.Point(0, 0);
            this.metroPanel1.Name = "metroPanel1";
            this.metroPanel1.Size = new System.Drawing.Size(424, 823);
            this.metroPanel1.TabIndex = 0;
            this.metroPanel1.VerticalScrollbarBarColor = true;
            this.metroPanel1.VerticalScrollbarHighlightOnWheel = false;
            this.metroPanel1.VerticalScrollbarSize = 10;
            // 
            // btnResetEvidenceNo
            // 
            this.btnResetEvidenceNo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnResetEvidenceNo.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnResetEvidenceNo.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnResetEvidenceNo.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnResetEvidenceNo.ImageKey = "duplicate.png";
            this.btnResetEvidenceNo.ImageList = this.imageList1;
            this.btnResetEvidenceNo.Location = new System.Drawing.Point(3, 107);
            this.btnResetEvidenceNo.Name = "btnResetEvidenceNo";
            this.btnResetEvidenceNo.Size = new System.Drawing.Size(418, 46);
            this.btnResetEvidenceNo.TabIndex = 22;
            this.btnResetEvidenceNo.Text = "Update All Evidence No.";
            this.btnResetEvidenceNo.UseVisualStyleBackColor = true;
            this.btnResetEvidenceNo.Click += new System.EventHandler(this.btnResetEvidenceNo_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "analysis.png");
            this.imageList1.Images.SetKeyName(1, "documentation.png");
            this.imageList1.Images.SetKeyName(2, "Marcus-Roberto-Google-Play-Google-Translate.512.png");
            this.imageList1.Images.SetKeyName(3, "settings.png");
            this.imageList1.Images.SetKeyName(4, "generative-image.png");
            this.imageList1.Images.SetKeyName(5, "duplicate.png");
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.BackColor = System.Drawing.Color.Transparent;
            this.groupBox2.Controls.Add(this.cbScalePercent);
            this.groupBox2.Controls.Add(this.cbAutoFixWidth);
            this.groupBox2.Controls.Add(this.numScalePercent);
            this.groupBox2.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold);
            this.groupBox2.Location = new System.Drawing.Point(4, 159);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(417, 57);
            this.groupBox2.TabIndex = 21;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Image Settings";
            // 
            // cbScalePercent
            // 
            this.cbScalePercent.AutoSize = true;
            this.cbScalePercent.Location = new System.Drawing.Point(113, 27);
            this.cbScalePercent.Name = "cbScalePercent";
            this.cbScalePercent.Size = new System.Drawing.Size(109, 15);
            this.cbScalePercent.TabIndex = 23;
            this.cbScalePercent.Text = "By Scale Percent";
            this.cbScalePercent.UseVisualStyleBackColor = true;
            // 
            // cbAutoFixWidth
            // 
            this.cbAutoFixWidth.AutoSize = true;
            this.cbAutoFixWidth.Checked = true;
            this.cbAutoFixWidth.Location = new System.Drawing.Point(6, 27);
            this.cbAutoFixWidth.Name = "cbAutoFixWidth";
            this.cbAutoFixWidth.Size = new System.Drawing.Size(101, 15);
            this.cbAutoFixWidth.TabIndex = 22;
            this.cbAutoFixWidth.TabStop = true;
            this.cbAutoFixWidth.Text = "Auto Fix Width";
            this.cbAutoFixWidth.UseVisualStyleBackColor = true;
            this.cbAutoFixWidth.CheckedChanged += new System.EventHandler(this.cbAutoFixWidth_CheckedChanged);
            // 
            // numScalePercent
            // 
            this.numScalePercent.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.numScalePercent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.numScalePercent.Enabled = false;
            this.numScalePercent.Location = new System.Drawing.Point(228, 23);
            this.numScalePercent.Name = "numScalePercent";
            this.numScalePercent.Size = new System.Drawing.Size(183, 23);
            this.numScalePercent.TabIndex = 3;
            this.numScalePercent.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.numScalePercent.Value = new decimal(new int[] {
            90,
            0,
            0,
            0});
            // 
            // btnFormatImages
            // 
            this.btnFormatImages.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFormatImages.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFormatImages.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnFormatImages.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnFormatImages.ImageKey = "generative-image.png";
            this.btnFormatImages.ImageList = this.imageList1;
            this.btnFormatImages.Location = new System.Drawing.Point(4, 348);
            this.btnFormatImages.Name = "btnFormatImages";
            this.btnFormatImages.Size = new System.Drawing.Size(417, 46);
            this.btnFormatImages.TabIndex = 21;
            this.btnFormatImages.Text = "Format Images";
            this.btnFormatImages.UseVisualStyleBackColor = true;
            this.btnFormatImages.Click += new System.EventHandler(this.btnFormatImages_Click);
            // 
            // toolStrip1
            // 
            this.toolStrip1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripFilePath});
            this.toolStrip1.Location = new System.Drawing.Point(0, 798);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(424, 25);
            this.toolStrip1.TabIndex = 20;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripFilePath
            // 
            this.toolStripFilePath.Font = new System.Drawing.Font("Segoe UI Semibold", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStripFilePath.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.toolStripFilePath.Name = "toolStripFilePath";
            this.toolStripFilePath.Size = new System.Drawing.Size(38, 22);
            this.toolStripFilePath.Text = "Book1";
            this.toolStripFilePath.Click += new System.EventHandler(this.toolStripFilePath_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.BackColor = System.Drawing.Color.Transparent;
            this.groupBox1.Controls.Add(this.chkInsertOnNewPage);
            this.groupBox1.Controls.Add(this.btnInsertPictures);
            this.groupBox1.Controls.Add(this.txtImagePath);
            this.groupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold);
            this.groupBox1.Location = new System.Drawing.Point(4, 222);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(417, 120);
            this.groupBox1.TabIndex = 19;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Insert multiple images";
            // 
            // chkInsertOnNewPage
            // 
            this.chkInsertOnNewPage.AutoSize = true;
            this.chkInsertOnNewPage.Checked = true;
            this.chkInsertOnNewPage.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkInsertOnNewPage.Location = new System.Drawing.Point(6, 96);
            this.chkInsertOnNewPage.Name = "chkInsertOnNewPage";
            this.chkInsertOnNewPage.Size = new System.Drawing.Size(127, 15);
            this.chkInsertOnNewPage.TabIndex = 4;
            this.chkInsertOnNewPage.Text = "Insert On New Page";
            this.chkInsertOnNewPage.UseVisualStyleBackColor = true;
            // 
            // btnInsertPictures
            // 
            this.btnInsertPictures.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnInsertPictures.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnInsertPictures.Image = global::ExcelCustomAddin.Properties.Resources.pictures1;
            this.btnInsertPictures.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnInsertPictures.Location = new System.Drawing.Point(6, 51);
            this.btnInsertPictures.Name = "btnInsertPictures";
            this.btnInsertPictures.Size = new System.Drawing.Size(405, 39);
            this.btnInsertPictures.TabIndex = 2;
            this.btnInsertPictures.Text = "Insert images";
            this.btnInsertPictures.UseVisualStyleBackColor = true;
            this.btnInsertPictures.Click += new System.EventHandler(this.btnInsertPictures_Click);
            // 
            // txtImagePath
            // 
            this.txtImagePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtImagePath.Location = new System.Drawing.Point(6, 22);
            this.txtImagePath.Multiline = true;
            this.txtImagePath.Name = "txtImagePath";
            this.txtImagePath.PromptText = "Image Path";
            this.txtImagePath.Size = new System.Drawing.Size(405, 23);
            this.txtImagePath.TabIndex = 1;
            this.txtImagePath.Text = "C:\\Images";
            // 
            // btnCreateEvidence
            // 
            this.btnCreateEvidence.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCreateEvidence.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateEvidence.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnCreateEvidence.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCreateEvidence.ImageKey = "documentation.png";
            this.btnCreateEvidence.ImageList = this.imageList1;
            this.btnCreateEvidence.Location = new System.Drawing.Point(3, 55);
            this.btnCreateEvidence.Name = "btnCreateEvidence";
            this.btnCreateEvidence.Size = new System.Drawing.Size(418, 46);
            this.btnCreateEvidence.TabIndex = 18;
            this.btnCreateEvidence.Text = "Create Evidence Sheet";
            this.btnCreateEvidence.UseVisualStyleBackColor = true;
            this.btnCreateEvidence.Click += new System.EventHandler(this.btnCreateEvidence_Click);
            // 
            // btnFormatDocument
            // 
            this.btnFormatDocument.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnFormatDocument.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFormatDocument.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnFormatDocument.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnFormatDocument.ImageKey = "analysis.png";
            this.btnFormatDocument.ImageList = this.imageList1;
            this.btnFormatDocument.Location = new System.Drawing.Point(3, 3);
            this.btnFormatDocument.Name = "btnFormatDocument";
            this.btnFormatDocument.Size = new System.Drawing.Size(418, 46);
            this.btnFormatDocument.TabIndex = 17;
            this.btnFormatDocument.Text = "Format Document";
            this.btnFormatDocument.UseVisualStyleBackColor = true;
            this.btnFormatDocument.Click += new System.EventHandler(this.btnFormatDocument_Click);
            // 
            // listofSheet
            // 
            this.listofSheet.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listofSheet.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.listofSheet.ContextMenuStrip = this.contextMenuStrip1;
            this.listofSheet.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listofSheet.FullRowSelect = true;
            this.listofSheet.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.listofSheet.HideSelection = false;
            this.listofSheet.Location = new System.Drawing.Point(3, 400);
            this.listofSheet.Name = "listofSheet";
            this.listofSheet.Size = new System.Drawing.Size(418, 395);
            this.listofSheet.TabIndex = 13;
            this.listofSheet.UseCompatibleStateImageBehavior = false;
            this.listofSheet.View = System.Windows.Forms.View.Details;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnChangeSheetName,
            this.btnPinSheet});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(188, 48);
            // 
            // btnChangeSheetName
            // 
            this.btnChangeSheetName.Image = global::ExcelCustomAddin.Properties.Resources.rename;
            this.btnChangeSheetName.Name = "btnChangeSheetName";
            this.btnChangeSheetName.Size = new System.Drawing.Size(187, 22);
            this.btnChangeSheetName.Text = "Change sheet\'s name";
            this.btnChangeSheetName.Click += new System.EventHandler(this.btnChangeSheetName_Click);
            // 
            // btnPinSheet
            // 
            this.btnPinSheet.Image = global::ExcelCustomAddin.Properties.Resources.pin;
            this.btnPinSheet.Name = "btnPinSheet";
            this.btnPinSheet.Size = new System.Drawing.Size(187, 22);
            this.btnPinSheet.Text = "Pin/Unpin Sheet";
            this.btnPinSheet.Click += new System.EventHandler(this.btnPinSheet_Click);
            // 
            // btnInsertMultipleImages
            // 
            this.btnInsertMultipleImages.Location = new System.Drawing.Point(0, 0);
            this.btnInsertMultipleImages.Name = "btnInsertMultipleImages";
            this.btnInsertMultipleImages.Size = new System.Drawing.Size(75, 23);
            this.btnInsertMultipleImages.TabIndex = 0;
            // 
            // ActionPanelControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.metroPanel1);
            this.Name = "ActionPanelControl";
            this.Size = new System.Drawing.Size(424, 823);
            this.metroPanel1.ResumeLayout(false);
            this.metroPanel1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numScalePercent)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private MetroFramework.Controls.MetroPanel metroPanel1;
        public System.Windows.Forms.ListView listofSheet;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem btnChangeSheetName;
        private System.Windows.Forms.ToolStripMenuItem btnPinSheet;
        private System.Windows.Forms.Button btnFormatDocument;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Button btnCreateEvidence;
        private System.Windows.Forms.Button btnInsertMultipleImages;
        private System.Windows.Forms.GroupBox groupBox1;
        public MetroFramework.Controls.MetroTextBox txtImagePath;
        private System.Windows.Forms.Button btnInsertPictures;
        public System.Windows.Forms.NumericUpDown numScalePercent;
        private System.Windows.Forms.ToolStrip toolStrip1;
        public System.Windows.Forms.ToolStripLabel toolStripFilePath;
        public MetroFramework.Controls.MetroCheckBox chkInsertOnNewPage;
        private System.Windows.Forms.Button btnFormatImages;
        private System.Windows.Forms.GroupBox groupBox2;
        public MetroFramework.Controls.MetroRadioButton cbAutoFixWidth;
        public MetroFramework.Controls.MetroRadioButton cbScalePercent;
        private System.Windows.Forms.Button btnResetEvidenceNo;
    }
}
