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
            this.metroPanel1 = new MetroFramework.Controls.MetroPanel();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.btnSheetTranslate = new MetroFramework.Controls.MetroButton();
            this.btnTranslateSelectedText = new MetroFramework.Controls.MetroButton();
            this.progressBar = new MetroFramework.Controls.MetroProgressBar();
            this.txtSourceText = new MetroFramework.Controls.MetroTextBox();
            this.txtDesText = new MetroFramework.Controls.MetroTextBox();
            this.metroPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // metroPanel1
            // 
            this.metroPanel1.Controls.Add(this.splitContainer1);
            this.metroPanel1.Controls.Add(this.btnSheetTranslate);
            this.metroPanel1.Controls.Add(this.btnTranslateSelectedText);
            this.metroPanel1.Controls.Add(this.progressBar);
            this.metroPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.metroPanel1.HorizontalScrollbarBarColor = true;
            this.metroPanel1.HorizontalScrollbarHighlightOnWheel = false;
            this.metroPanel1.HorizontalScrollbarSize = 10;
            this.metroPanel1.Location = new System.Drawing.Point(0, 0);
            this.metroPanel1.Name = "metroPanel1";
            this.metroPanel1.Size = new System.Drawing.Size(282, 823);
            this.metroPanel1.TabIndex = 13;
            this.metroPanel1.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.metroPanel1.VerticalScrollbarBarColor = true;
            this.metroPanel1.VerticalScrollbarHighlightOnWheel = false;
            this.metroPanel1.VerticalScrollbarSize = 10;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.txtSourceText);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.txtDesText);
            this.splitContainer1.Size = new System.Drawing.Size(282, 768);
            this.splitContainer1.SplitterDistance = 365;
            this.splitContainer1.TabIndex = 17;
            // 
            // btnSheetTranslate
            // 
            this.btnSheetTranslate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSheetTranslate.Location = new System.Drawing.Point(88, 774);
            this.btnSheetTranslate.Name = "btnSheetTranslate";
            this.btnSheetTranslate.Size = new System.Drawing.Size(194, 38);
            this.btnSheetTranslate.TabIndex = 16;
            this.btnSheetTranslate.Text = "SHEET TRANSLATE";
            this.btnSheetTranslate.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.btnSheetTranslate.Click += new System.EventHandler(this.btnSheetTranslate_Click);
            // 
            // btnTranslateSelectedText
            // 
            this.btnTranslateSelectedText.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnTranslateSelectedText.Location = new System.Drawing.Point(0, 773);
            this.btnTranslateSelectedText.Name = "btnTranslateSelectedText";
            this.btnTranslateSelectedText.Size = new System.Drawing.Size(82, 38);
            this.btnTranslateSelectedText.TabIndex = 15;
            this.btnTranslateSelectedText.Text = "TRANSLATE";
            this.btnTranslateSelectedText.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.btnTranslateSelectedText.Click += new System.EventHandler(this.btnTranslateSelectedText_Click);
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(0, 813);
            this.progressBar.Name = "progressBar";
            this.progressBar.ProgressBarStyle = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar.Size = new System.Drawing.Size(282, 10);
            this.progressBar.TabIndex = 14;
            this.progressBar.Theme = MetroFramework.MetroThemeStyle.Dark;
            this.progressBar.Visible = false;
            // 
            // txtSourceText
            // 
            this.txtSourceText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtSourceText.Location = new System.Drawing.Point(0, 0);
            this.txtSourceText.Multiline = true;
            this.txtSourceText.Name = "txtSourceText";
            this.txtSourceText.Size = new System.Drawing.Size(282, 365);
            this.txtSourceText.TabIndex = 13;
            this.txtSourceText.Theme = MetroFramework.MetroThemeStyle.Dark;
            // 
            // txtDesText
            // 
            this.txtDesText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtDesText.Location = new System.Drawing.Point(0, 0);
            this.txtDesText.Multiline = true;
            this.txtDesText.Name = "txtDesText";
            this.txtDesText.Size = new System.Drawing.Size(282, 399);
            this.txtDesText.TabIndex = 14;
            this.txtDesText.Theme = MetroFramework.MetroThemeStyle.Dark;
            // 
            // ActionPanelControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.metroPanel1);
            this.Name = "ActionPanelControl";
            this.Size = new System.Drawing.Size(282, 823);
            this.metroPanel1.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private MetroFramework.Controls.MetroPanel metroPanel1;
        public MetroFramework.Controls.MetroProgressBar progressBar;
        public MetroFramework.Controls.MetroButton btnSheetTranslate;
        public MetroFramework.Controls.MetroButton btnTranslateSelectedText;
        private System.Windows.Forms.SplitContainer splitContainer1;
        public MetroFramework.Controls.MetroTextBox txtSourceText;
        public MetroFramework.Controls.MetroTextBox txtDesText;
    }
}
