
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
            this.txtSourceText = new System.Windows.Forms.RichTextBox();
            this.btnTranslate = new System.Windows.Forms.Button();
            this.txtDesText = new System.Windows.Forms.RichTextBox();
            this.listOfSheet = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtSourceText
            // 
            this.txtSourceText.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSourceText.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtSourceText.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSourceText.Location = new System.Drawing.Point(0, 0);
            this.txtSourceText.Name = "txtSourceText";
            this.txtSourceText.Size = new System.Drawing.Size(402, 159);
            this.txtSourceText.TabIndex = 0;
            this.txtSourceText.Text = "";
            // 
            // btnTranslate
            // 
            this.btnTranslate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnTranslate.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnTranslate.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnTranslate.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnTranslate.Location = new System.Drawing.Point(240, 165);
            this.btnTranslate.Name = "btnTranslate";
            this.btnTranslate.Size = new System.Drawing.Size(162, 28);
            this.btnTranslate.TabIndex = 1;
            this.btnTranslate.Text = "TRANSLATE";
            this.btnTranslate.UseVisualStyleBackColor = false;
            this.btnTranslate.Click += new System.EventHandler(this.btnTranslate_Click);
            // 
            // txtDesText
            // 
            this.txtDesText.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDesText.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtDesText.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDesText.Location = new System.Drawing.Point(0, 199);
            this.txtDesText.Name = "txtDesText";
            this.txtDesText.Size = new System.Drawing.Size(402, 166);
            this.txtDesText.TabIndex = 2;
            this.txtDesText.Text = "";
            // 
            // listOfSheet
            // 
            this.listOfSheet.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listOfSheet.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listOfSheet.FormattingEnabled = true;
            this.listOfSheet.ItemHeight = 15;
            this.listOfSheet.Location = new System.Drawing.Point(0, 410);
            this.listOfSheet.Name = "listOfSheet";
            this.listOfSheet.Size = new System.Drawing.Size(402, 394);
            this.listOfSheet.TabIndex = 3;
            this.listOfSheet.SelectedIndexChanged += new System.EventHandler(this.listOfSheet_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(0, 386);
            this.label1.Margin = new System.Windows.Forms.Padding(0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 15);
            this.label1.TabIndex = 4;
            this.label1.Text = "Danh sách sheet";
            // 
            // ActionPanelControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listOfSheet);
            this.Controls.Add(this.txtDesText);
            this.Controls.Add(this.btnTranslate);
            this.Controls.Add(this.txtSourceText);
            this.Name = "ActionPanelControl";
            this.Size = new System.Drawing.Size(402, 823);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnTranslate;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.ListBox listOfSheet;
        public System.Windows.Forms.RichTextBox txtSourceText;
        public System.Windows.Forms.RichTextBox txtDesText;
    }
}
