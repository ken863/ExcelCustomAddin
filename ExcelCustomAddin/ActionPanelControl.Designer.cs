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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.numScalePercent = new System.Windows.Forms.NumericUpDown();
            this.btnInsertPictures = new System.Windows.Forms.Button();
            this.txtImagePath = new MetroFramework.Controls.MetroTextBox();
            this.metroLabel1 = new MetroFramework.Controls.MetroLabel();
            this.btnCreateEvidence = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.btnFormatDocument = new System.Windows.Forms.Button();
            this.txtSheetListLabel = new MetroFramework.Controls.MetroLabel();
            this.listofSheet = new System.Windows.Forms.ListView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.btnChangeSheetName = new System.Windows.Forms.ToolStripMenuItem();
            this.btnPinSheet = new System.Windows.Forms.ToolStripMenuItem();
            this.btnInsertMultipleImages = new System.Windows.Forms.Button();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripFilePath = new System.Windows.Forms.ToolStripLabel();
            this.metroPanel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numScalePercent)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // metroPanel1
            // 
            this.metroPanel1.Controls.Add(this.toolStrip1);
            this.metroPanel1.Controls.Add(this.groupBox1);
            this.metroPanel1.Controls.Add(this.btnCreateEvidence);
            this.metroPanel1.Controls.Add(this.btnFormatDocument);
            this.metroPanel1.Controls.Add(this.txtSheetListLabel);
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
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.BackColor = System.Drawing.Color.Transparent;
            this.groupBox1.Controls.Add(this.numScalePercent);
            this.groupBox1.Controls.Add(this.btnInsertPictures);
            this.groupBox1.Controls.Add(this.txtImagePath);
            this.groupBox1.Controls.Add(this.metroLabel1);
            this.groupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold);
            this.groupBox1.Location = new System.Drawing.Point(4, 108);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(417, 96);
            this.groupBox1.TabIndex = 19;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Insert multiple images";
            // 
            // numScalePercent
            // 
            this.numScalePercent.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.numScalePercent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.numScalePercent.Location = new System.Drawing.Point(366, 22);
            this.numScalePercent.Name = "numScalePercent";
            this.numScalePercent.Size = new System.Drawing.Size(45, 23);
            this.numScalePercent.TabIndex = 3;
            this.numScalePercent.Value = new decimal(new int[] {
            90,
            0,
            0,
            0});
            // 
            // btnInsertPictures
            // 
            this.btnInsertPictures.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnInsertPictures.Image = global::ExcelCustomAddin.Properties.Resources.pictures1;
            this.btnInsertPictures.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnInsertPictures.Location = new System.Drawing.Point(85, 51);
            this.btnInsertPictures.Name = "btnInsertPictures";
            this.btnInsertPictures.Size = new System.Drawing.Size(326, 39);
            this.btnInsertPictures.TabIndex = 2;
            this.btnInsertPictures.Text = "Insert images";
            this.btnInsertPictures.UseVisualStyleBackColor = true;
            this.btnInsertPictures.Click += new System.EventHandler(this.btnInsertPictures_Click);
            // 
            // txtImagePath
            // 
            this.txtImagePath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtImagePath.Location = new System.Drawing.Point(85, 22);
            this.txtImagePath.Multiline = true;
            this.txtImagePath.Name = "txtImagePath";
            this.txtImagePath.Size = new System.Drawing.Size(275, 23);
            this.txtImagePath.TabIndex = 1;
            this.txtImagePath.Text = "C:\\Images";
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.FontSize = MetroFramework.MetroLabelSize.Small;
            this.metroLabel1.FontWeight = MetroFramework.MetroLabelWeight.Regular;
            this.metroLabel1.Location = new System.Drawing.Point(6, 25);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(72, 15);
            this.metroLabel1.TabIndex = 0;
            this.metroLabel1.Text = "Images Path";
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
            this.btnCreateEvidence.Text = "Create Evidence";
            this.btnCreateEvidence.UseVisualStyleBackColor = true;
            this.btnCreateEvidence.Click += new System.EventHandler(this.btnCreateEvidence_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "analysis.png");
            this.imageList1.Images.SetKeyName(1, "documentation.png");
            this.imageList1.Images.SetKeyName(2, "Marcus-Roberto-Google-Play-Google-Translate.512.png");
            this.imageList1.Images.SetKeyName(3, "settings.png");
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
            // txtSheetListLabel
            // 
            this.txtSheetListLabel.AutoSize = true;
            this.txtSheetListLabel.FontSize = MetroFramework.MetroLabelSize.Small;
            this.txtSheetListLabel.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            this.txtSheetListLabel.Location = new System.Drawing.Point(3, 207);
            this.txtSheetListLabel.Name = "txtSheetListLabel";
            this.txtSheetListLabel.Size = new System.Drawing.Size(69, 15);
            this.txtSheetListLabel.TabIndex = 16;
            this.txtSheetListLabel.Text = "SHEET LIST";
            // 
            // listofSheet
            // 
            this.listofSheet.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listofSheet.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.listofSheet.ContextMenuStrip = this.contextMenuStrip1;
            this.listofSheet.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listofSheet.FullRowSelect = true;
            this.listofSheet.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None;
            this.listofSheet.HideSelection = false;
            this.listofSheet.Location = new System.Drawing.Point(3, 225);
            this.listofSheet.Name = "listofSheet";
            this.listofSheet.Size = new System.Drawing.Size(418, 570);
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
            this.toolStripFilePath.Name = "toolStripFilePath";
            this.toolStripFilePath.Size = new System.Drawing.Size(0, 22);
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
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numScalePercent)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private MetroFramework.Controls.MetroPanel metroPanel1;
        public System.Windows.Forms.ListView listofSheet;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem btnChangeSheetName;
        private System.Windows.Forms.ToolStripMenuItem btnPinSheet;
        private MetroFramework.Controls.MetroLabel txtSheetListLabel;
        private System.Windows.Forms.Button btnFormatDocument;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Button btnCreateEvidence;
        private System.Windows.Forms.Button btnInsertMultipleImages;
        private System.Windows.Forms.GroupBox groupBox1;
        private MetroFramework.Controls.MetroLabel metroLabel1;
        public MetroFramework.Controls.MetroTextBox txtImagePath;
        private System.Windows.Forms.Button btnInsertPictures;
        public System.Windows.Forms.NumericUpDown numScalePercent;
        private System.Windows.Forms.ToolStrip toolStrip1;
        public System.Windows.Forms.ToolStripLabel toolStripFilePath;
    }
}
