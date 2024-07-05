
using System;

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
            this.buttonTranslate = new System.Windows.Forms.Button();
            this.txtDesText = new System.Windows.Forms.RichTextBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.bgwTranslate = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // txtSourceText
            // 
            this.txtSourceText.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSourceText.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtSourceText.Font = new System.Drawing.Font("Meiryo UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSourceText.Location = new System.Drawing.Point(0, 0);
            this.txtSourceText.Name = "txtSourceText";
            this.txtSourceText.Size = new System.Drawing.Size(402, 285);
            this.txtSourceText.TabIndex = 4;
            this.txtSourceText.Text = "";
            // 
            // buttonTranslate
            // 
            this.buttonTranslate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonTranslate.Font = new System.Drawing.Font("Meiryo UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonTranslate.ForeColor = System.Drawing.Color.Black;
            this.buttonTranslate.Location = new System.Drawing.Point(0, 291);
            this.buttonTranslate.Name = "buttonTranslate";
            this.buttonTranslate.Size = new System.Drawing.Size(402, 34);
            this.buttonTranslate.TabIndex = 7;
            this.buttonTranslate.Text = "TRANSLATE";
            this.buttonTranslate.UseVisualStyleBackColor = true;
            this.buttonTranslate.Click += new System.EventHandler(this.ButtonTranslate_Click);
            // 
            // txtDesText
            // 
            this.txtDesText.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtDesText.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txtDesText.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDesText.Location = new System.Drawing.Point(0, 331);
            this.txtDesText.Name = "txtDesText";
            this.txtDesText.ReadOnly = true;
            this.txtDesText.Size = new System.Drawing.Size(402, 476);
            this.txtDesText.TabIndex = 8;
            this.txtDesText.Text = "";
            // 
            // progressBar
            // 
            this.progressBar.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar.Location = new System.Drawing.Point(0, 813);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(402, 10);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar.TabIndex = 9;
            this.progressBar.Visible = false;
            // 
            // bgwTranslate
            // 
            this.bgwTranslate.WorkerReportsProgress = true;
            this.bgwTranslate.DoWork += new System.ComponentModel.DoWorkEventHandler(this.bgwTranslate_DoWork);
            // 
            // ActionPanelControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.txtDesText);
            this.Controls.Add(this.buttonTranslate);
            this.Controls.Add(this.txtSourceText);
            this.Name = "ActionPanelControl";
            this.Size = new System.Drawing.Size(402, 823);
            this.ResumeLayout(false);

        }

        #endregion
        public System.Windows.Forms.RichTextBox txtSourceText;
        private System.Windows.Forms.Button buttonTranslate;
        public System.Windows.Forms.RichTextBox txtDesText;
        public System.Windows.Forms.ProgressBar progressBar;
        public System.ComponentModel.BackgroundWorker bgwTranslate;
    }
}
