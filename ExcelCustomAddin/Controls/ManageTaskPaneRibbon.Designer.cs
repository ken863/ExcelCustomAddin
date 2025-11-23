namespace ExcelCustomAddin
{
    partial class ManageTaskPaneRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ManageTaskPaneRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnTranslate = this.Factory.CreateRibbonButton();
            this.btnSheetConfigManager = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TOOLS";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnTranslate);
            this.group1.Items.Add(this.btnSheetConfigManager);
            this.group1.Label = "Tools";
            this.group1.Name = "group1";
            // 
            // btnTranslate
            // 
            this.btnTranslate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTranslate.Image = global::ExcelCustomAddin.Properties.Resources.suitcase;
            this.btnTranslate.Label = "TOOLS";
            this.btnTranslate.Name = "btnTranslate";
            this.btnTranslate.ShowImage = true;
            this.btnTranslate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTranslate_Click);
            // 
            // btnSheetConfigManager
            // 
            this.btnSheetConfigManager.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSheetConfigManager.Image = global::ExcelCustomAddin.Properties.Resources.settings;
            this.btnSheetConfigManager.Label = "CONFIG";
            this.btnSheetConfigManager.Name = "btnSheetConfigManager";
            this.btnSheetConfigManager.ShowImage = true;
            this.btnSheetConfigManager.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSheetConfigManager_Click);
            // 
            // ManageTaskPaneRibbon
            // 
            this.Name = "ManageTaskPaneRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTranslate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSheetConfigManager;
    }

    partial class ThisRibbonCollection
    {
        internal ManageTaskPaneRibbon ManageTaskPaneRibbon
        {
            get { return this.GetRibbon<ManageTaskPaneRibbon>(); }
        }
    }
}
