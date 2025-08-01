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
            this.listofSheet = new System.Windows.Forms.ListBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.btnChangeSheetName = new System.Windows.Forms.ToolStripMenuItem();
            this.metroLabel1 = new MetroFramework.Controls.MetroLabel();
            this.btnFormatDocument = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.btnCreateEvidence = new System.Windows.Forms.Button();
            this.metroPanel1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // metroPanel1
            // 
            this.metroPanel1.Controls.Add(this.btnCreateEvidence);
            this.metroPanel1.Controls.Add(this.btnFormatDocument);
            this.metroPanel1.Controls.Add(this.metroLabel1);
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
            // listofSheet
            // 
            this.listofSheet.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listofSheet.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.listofSheet.ContextMenuStrip = this.contextMenuStrip1;
            this.listofSheet.Font = new System.Drawing.Font("Segoe UI Semibold", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listofSheet.FormattingEnabled = true;
            this.listofSheet.ItemHeight = 15;
            this.listofSheet.Location = new System.Drawing.Point(3, 126);
            this.listofSheet.Name = "listofSheet";
            this.listofSheet.Size = new System.Drawing.Size(418, 690);
            this.listofSheet.TabIndex = 13;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnChangeSheetName});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(188, 26);
            // 
            // btnChangeSheetName
            // 
            this.btnChangeSheetName.Name = "btnChangeSheetName";
            this.btnChangeSheetName.Size = new System.Drawing.Size(187, 22);
            this.btnChangeSheetName.Text = "Change sheet\'s name";
            this.btnChangeSheetName.Click += new System.EventHandler(this.btnChangeSheetName_Click);
            // 
            // metroLabel1
            // 
            this.metroLabel1.AutoSize = true;
            this.metroLabel1.FontSize = MetroFramework.MetroLabelSize.Small;
            this.metroLabel1.FontWeight = MetroFramework.MetroLabelWeight.Bold;
            this.metroLabel1.Location = new System.Drawing.Point(0, 104);
            this.metroLabel1.Name = "metroLabel1";
            this.metroLabel1.Size = new System.Drawing.Size(69, 15);
            this.metroLabel1.TabIndex = 16;
            this.metroLabel1.Text = "SHEET LIST";
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
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "analysis.png");
            this.imageList1.Images.SetKeyName(1, "documentation.png");
            this.imageList1.Images.SetKeyName(2, "Marcus-Roberto-Google-Play-Google-Translate.512.png");
            this.imageList1.Images.SetKeyName(3, "settings.png");
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
            // ActionPanelControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.metroPanel1);
            this.Name = "ActionPanelControl";
            this.Size = new System.Drawing.Size(424, 823);
            this.metroPanel1.ResumeLayout(false);
            this.metroPanel1.PerformLayout();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private MetroFramework.Controls.MetroPanel metroPanel1;
        public System.Windows.Forms.ListBox listofSheet;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem btnChangeSheetName;
        private MetroFramework.Controls.MetroLabel metroLabel1;
        private System.Windows.Forms.Button btnFormatDocument;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Button btnCreateEvidence;
    }
}
