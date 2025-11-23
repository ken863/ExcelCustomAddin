using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelCustomAddin.Controls
{
    public partial class SheetConfigManagerForm : Form
    {
        private bool _isDirty = false;

        public SheetConfigManagerForm()
        {
            InitializeComponent();
            // LoadConfiguration() will be called in Shown event
        }

        private void InitializeComponent()
        {
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabSheets = new System.Windows.Forms.TabPage();
            this.tabGeneralSettings = new System.Windows.Forms.TabPage();
            this.tabLoggingSettings = new System.Windows.Forms.TabPage();

            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnApply = new System.Windows.Forms.Button();

            // Initialize Sheets tab
            this.lblSheets = new System.Windows.Forms.Label();
            this.dgvSheets = new System.Windows.Forms.DataGridView();
            this.btnAddSheet = new System.Windows.Forms.Button();
            this.btnRemoveSheet = new System.Windows.Forms.Button();

            // Tab Control
            this.tabControl.SuspendLayout();
            this.tabControl.Controls.Add(this.tabSheets);
            this.tabControl.Controls.Add(this.tabGeneralSettings);
            this.tabControl.Controls.Add(this.tabLoggingSettings);
            this.tabControl.Location = new System.Drawing.Point(12, 12);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(760, 480);
            this.tabControl.TabIndex = 0;

            // Sheets Tab
            this.tabSheets.Controls.Add(this.lblSheets);
            this.tabSheets.Controls.Add(this.dgvSheets);
            this.tabSheets.Controls.Add(this.btnAddSheet);
            this.tabSheets.Controls.Add(this.btnRemoveSheet);
            this.tabSheets.Location = new System.Drawing.Point(4, 22);
            this.tabSheets.Name = "tabSheets";
            this.tabSheets.Padding = new System.Windows.Forms.Padding(3);
            this.tabSheets.Size = new System.Drawing.Size(752, 454);
            this.tabSheets.TabIndex = 0;
            this.tabSheets.Text = "Sheets";
            this.tabSheets.UseVisualStyleBackColor = true;

            // General Settings Tab
            this.tabGeneralSettings.Location = new System.Drawing.Point(4, 22);
            this.tabGeneralSettings.Name = "tabGeneralSettings";
            this.tabGeneralSettings.Padding = new System.Windows.Forms.Padding(3);
            this.tabGeneralSettings.Size = new System.Drawing.Size(752, 454);
            this.tabGeneralSettings.TabIndex = 1;
            this.tabGeneralSettings.Text = "General Settings";
            this.tabGeneralSettings.UseVisualStyleBackColor = true;

            // Logging Settings Tab
            this.tabLoggingSettings.Location = new System.Drawing.Point(4, 22);
            this.tabLoggingSettings.Name = "tabLoggingSettings";
            this.tabLoggingSettings.Padding = new System.Windows.Forms.Padding(3);
            this.tabLoggingSettings.Size = new System.Drawing.Size(752, 454);
            this.tabLoggingSettings.TabIndex = 2;
            this.tabLoggingSettings.Text = "Logging Settings";
            this.tabLoggingSettings.UseVisualStyleBackColor = true;

            // Buttons
            this.btnSave.Location = new System.Drawing.Point(608, 500);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);

            this.btnCancel.Location = new System.Drawing.Point(689, 500);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);

            this.btnApply.Location = new System.Drawing.Point(527, 500);
            this.btnApply.Name = "btnApply";
            this.btnApply.Size = new System.Drawing.Size(75, 23);
            this.btnApply.TabIndex = 3;
            this.btnApply.Text = "Apply";
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);

            // Form properties
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 535);
            this.Controls.Add(this.btnApply);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.tabControl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SheetConfigManagerForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Sheet Configuration Manager";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SheetConfigManagerForm_FormClosing);
            this.Shown += new System.EventHandler(this.SheetConfigManagerForm_Shown);
            this.tabControl.ResumeLayout(false);
            this.ResumeLayout(false);

            // Setup controls
            SetupSheetsTab();
            SetupGeneralSettingsTab();
            SetupLoggingSettingsTab();
        }

        private void SetupSheetsTab()
        {
            // Label
            this.lblSheets.AutoSize = true;
            this.lblSheets.Location = new System.Drawing.Point(6, 6);
            this.lblSheets.Name = "lblSheets";
            this.lblSheets.Size = new System.Drawing.Size(80, 13);
            this.lblSheets.TabIndex = 0;
            this.lblSheets.Text = "All Sheets:";

            // DataGridView
            this.dgvSheets.AllowUserToAddRows = false;
            this.dgvSheets.AllowUserToDeleteRows = false;
            this.dgvSheets.AllowUserToResizeColumns = true;
            this.dgvSheets.AllowUserToResizeRows = false;
            this.dgvSheets.AutoGenerateColumns = true; // Ensure columns are auto-generated
            this.dgvSheets.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSheets.Location = new System.Drawing.Point(6, 25);
            this.dgvSheets.Name = "dgvSheets";
            this.dgvSheets.ReadOnly = false;
            this.dgvSheets.RowHeadersVisible = false;
            this.dgvSheets.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvSheets.Size = new System.Drawing.Size(740, 380);
            this.dgvSheets.TabIndex = 1;
            this.dgvSheets.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvSheets_CellValueChanged);
            this.dgvSheets.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.dgvSheets_DataBindingComplete);

            // Add Button
            this.btnAddSheet.Location = new System.Drawing.Point(6, 415);
            this.btnAddSheet.Name = "btnAddSheet";
            this.btnAddSheet.Size = new System.Drawing.Size(75, 23);
            this.btnAddSheet.TabIndex = 2;
            this.btnAddSheet.Text = "Add";
            this.btnAddSheet.UseVisualStyleBackColor = true;
            this.btnAddSheet.Click += new System.EventHandler(this.btnAddSheet_Click);

            // Remove Button
            this.btnRemoveSheet.Location = new System.Drawing.Point(87, 415);
            this.btnRemoveSheet.Name = "btnRemoveSheet";
            this.btnRemoveSheet.Size = new System.Drawing.Size(75, 23);
            this.btnRemoveSheet.TabIndex = 3;
            this.btnRemoveSheet.Text = "Remove";
            this.btnRemoveSheet.UseVisualStyleBackColor = true;
            this.btnRemoveSheet.Click += new System.EventHandler(this.btnRemoveSheet_Click);
        }

        private void LoadConfiguration()
        {
            try
            {
                // Load All Sheets
                LoadAllSheets();

                // Load General Settings
                LoadGeneralSettings();

                // Load Logging Settings
                LoadLoggingSettings();

                _isDirty = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading configuration: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadAllSheets()
        {
            var allSheets = SheetConfigManager.GetAllSheetConfigs();
            
            // Debug logging
            System.Diagnostics.Debug.WriteLine($"Loading {allSheets.Count} sheets");
            foreach (var sheet in allSheets)
            {
                System.Diagnostics.Debug.WriteLine($"Sheet: {sheet.Name}, Prefix: {sheet.Prefix}");
            }
            
            var dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Prefix");
            dt.Columns.Add("ReferenceColumnHeader");
            dt.Columns.Add("NumberFormat");
            dt.Columns.Add("Description");
            dt.Columns.Add("IsHorizontal");

            foreach (var sheet in allSheets)
            {
                dt.Rows.Add(sheet.Name, sheet.Prefix, sheet.ReferenceColumnHeader, sheet.NumberFormat, sheet.Description, sheet.IsHorizontal);
                System.Diagnostics.Debug.WriteLine($"Added row: Name={sheet.Name}, Prefix={sheet.Prefix}, RefHeader={sheet.ReferenceColumnHeader}");
            }

            System.Diagnostics.Debug.WriteLine($"DataTable has {dt.Rows.Count} rows");
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                System.Diagnostics.Debug.WriteLine($"Row {i}: {dt.Rows[i]["Name"]}, {dt.Rows[i]["Prefix"]}");
            }

            // Clear existing data source first
            this.dgvSheets.DataSource = null;
            
            // Set new data source
            this.dgvSheets.DataSource = dt;
            
            // Force refresh and layout
            this.dgvSheets.Refresh();
            this.dgvSheets.Invalidate();
            this.dgvSheets.Update();
            this.tabSheets.Refresh();
            
            // Ensure columns are visible
            if (this.dgvSheets.Columns.Count > 0)
            {
                foreach (DataGridViewColumn col in this.dgvSheets.Columns)
                {
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    col.MinimumWidth = 50; // Ensure minimum width
                    col.Visible = true; // Ensure column is visible
                }
                
                // Specifically ensure ReferenceColumnHeader column is visible
                if (this.dgvSheets.Columns.Contains("ReferenceColumnHeader"))
                {
                    var refCol = this.dgvSheets.Columns["ReferenceColumnHeader"];
                    refCol.HeaderText = "Reference Column Header"; // More readable header
                    refCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    refCol.MinimumWidth = 120;
                    System.Diagnostics.Debug.WriteLine($"ReferenceColumnHeader column visible: {refCol.Visible}, width: {refCol.Width}");
                }
            }
            
            System.Diagnostics.Debug.WriteLine($"DataGridView has {this.dgvSheets.Rows.Count} rows after binding");
        }

        private void LoadGeneralSettings()
        {
            var generalConfig = SheetConfigManager.GetGeneralConfig();

            if (this.chkAutoFillCell != null)
                this.chkAutoFillCell.Checked = generalConfig.AutoFillCell;
            if (this.chkEnableDebugLog != null)
                this.chkEnableDebugLog.Checked = generalConfig.EnableDebugLog;
            if (this.numStartingNumber != null)
                this.numStartingNumber.Value = generalConfig.StartingNumber;
            if (this.txtPageBreakColumnName != null)
                this.txtPageBreakColumnName.Text = generalConfig.PageBreakColumnName;
            if (this.txtEvidenceFontName != null)
                this.txtEvidenceFontName.Text = generalConfig.EvidenceFontName;
            if (this.txtBackButtonFontName != null)
                this.txtBackButtonFontName.Text = generalConfig.BackButtonFontName;
            if (this.numPrintAreaLastRowIdx != null)
                this.numPrintAreaLastRowIdx.Value = generalConfig.PrintAreaLastRowIdx;
            if (this.numColumnWidth != null)
                this.numColumnWidth.Value = (decimal)generalConfig.ColumnWidth;
            if (this.numRowHeight != null)
                this.numRowHeight.Value = (decimal)generalConfig.RowHeight;
            if (this.numFontSize != null)
                this.numFontSize.Value = generalConfig.FontSize;
            if (this.cmbPageOrientation != null)
                this.cmbPageOrientation.Text = generalConfig.PageOrientation;
            if (this.cmbPaperSize != null)
                this.cmbPaperSize.Text = generalConfig.PaperSize;
            if (this.numZoom != null)
                this.numZoom.Value = generalConfig.Zoom;
            if (this.chkFitToPagesWide != null)
                this.chkFitToPagesWide.Checked = generalConfig.FitToPagesWide;
            if (this.chkFitToPagesTall != null)
                this.chkFitToPagesTall.Checked = generalConfig.FitToPagesTall;
            if (this.chkCenterHorizontally != null)
                this.chkCenterHorizontally.Checked = generalConfig.CenterHorizontally;
            if (this.numWindowZoom != null)
                this.numWindowZoom.Value = generalConfig.WindowZoom;
            if (this.cmbViewMode != null)
                this.cmbViewMode.Text = generalConfig.ViewMode;
        }

        private void LoadLoggingSettings()
        {
            var loggingConfig = SheetConfigManager.GetLoggingConfig();

            if (this.txtLogDirectory != null)
                this.txtLogDirectory.Text = loggingConfig.LogDirectory;
            if (this.chkEnableFileLogging != null)
                this.chkEnableFileLogging.Checked = loggingConfig.EnableFileLogging;
            if (this.chkEnableDebugOutput != null)
                this.chkEnableDebugOutput.Checked = loggingConfig.EnableDebugOutput;
            if (this.cmbLogLevel != null)
                this.cmbLogLevel.Text = loggingConfig.LogLevel;
            if (this.txtLogFileName != null)
                this.txtLogFileName.Text = loggingConfig.LogFileName;
        }

        private void SaveConfiguration()
        {
            try
            {
                // Save All Sheets
                SaveAllSheets();

                // Save General Settings
                SaveGeneralSettings();

                // Save Logging Settings
                SaveLoggingSettings();

                SheetConfigManager.SaveConfiguration();
                _isDirty = false;

                MessageBox.Show("Configuration saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving configuration: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveAllSheets()
        {
            var allSheets = new List<SheetConfigManager.SheetConfig>();
            var dt = (DataTable)this.dgvSheets.DataSource;

            if (dt != null)
            {
                foreach (DataRow row in dt.Rows)
                {
                    allSheets.Add(new SheetConfigManager.SheetConfig
                    {
                        Name = row["Name"]?.ToString() ?? "",
                        Prefix = row["Prefix"]?.ToString() ?? "",
                        ReferenceColumnHeader = row["ReferenceColumnHeader"]?.ToString() ?? "",
                        NumberFormat = row["NumberFormat"]?.ToString() ?? "D2",
                        Description = row["Description"]?.ToString() ?? "",
                        IsHorizontal = row["IsHorizontal"]?.ToString() ?? "False"
                    });
                }
            }

            // Update the static field
            SheetConfigManager._sheets = allSheets;
        }

        private void SaveGeneralSettings()
        {
            var generalConfig = SheetConfigManager.GetGeneralConfig();

            if (this.chkAutoFillCell != null)
                generalConfig.AutoFillCell = this.chkAutoFillCell.Checked;
            if (this.chkEnableDebugLog != null)
                generalConfig.EnableDebugLog = this.chkEnableDebugLog.Checked;
            if (this.numStartingNumber != null)
                generalConfig.StartingNumber = (int)this.numStartingNumber.Value;
            if (this.txtPageBreakColumnName != null)
                generalConfig.PageBreakColumnName = this.txtPageBreakColumnName.Text;
            if (this.txtEvidenceFontName != null)
                generalConfig.EvidenceFontName = this.txtEvidenceFontName.Text;
            if (this.txtBackButtonFontName != null)
                generalConfig.BackButtonFontName = this.txtBackButtonFontName.Text;
            if (this.numPrintAreaLastRowIdx != null)
                generalConfig.PrintAreaLastRowIdx = (int)this.numPrintAreaLastRowIdx.Value;
            if (this.numColumnWidth != null)
                generalConfig.ColumnWidth = (double)this.numColumnWidth.Value;
            if (this.numRowHeight != null)
                generalConfig.RowHeight = (double)this.numRowHeight.Value;
            if (this.numFontSize != null)
                generalConfig.FontSize = (int)this.numFontSize.Value;
            if (this.cmbPageOrientation != null)
                generalConfig.PageOrientation = this.cmbPageOrientation.Text;
            if (this.cmbPaperSize != null)
                generalConfig.PaperSize = this.cmbPaperSize.Text;
            if (this.numZoom != null)
                generalConfig.Zoom = (int)this.numZoom.Value;
            if (this.chkFitToPagesWide != null)
                generalConfig.FitToPagesWide = this.chkFitToPagesWide.Checked;
            if (this.chkFitToPagesTall != null)
                generalConfig.FitToPagesTall = this.chkFitToPagesTall.Checked;
            if (this.chkCenterHorizontally != null)
                generalConfig.CenterHorizontally = this.chkCenterHorizontally.Checked;
            if (this.numWindowZoom != null)
                generalConfig.WindowZoom = (int)this.numWindowZoom.Value;
            if (this.cmbViewMode != null)
                generalConfig.ViewMode = this.cmbViewMode.Text;
        }

        private void SaveLoggingSettings()
        {
            var loggingConfig = SheetConfigManager.GetLoggingConfig();

            loggingConfig.LogDirectory = this.txtLogDirectory.Text;
            loggingConfig.EnableFileLogging = this.chkEnableFileLogging.Checked;
            loggingConfig.EnableDebugOutput = this.chkEnableDebugOutput.Checked;
            loggingConfig.LogLevel = this.cmbLogLevel.Text;
            loggingConfig.LogFileName = this.txtLogFileName.Text;
        }

        private void Control_ValueChanged(object sender, EventArgs e)
        {
            _isDirty = true;
        }

        private void dgvSheets_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            _isDirty = true;
        }

        private void btnAddSheet_Click(object sender, EventArgs e)
        {
            // Add new row to DataGridView
            var dt = (DataTable)this.dgvSheets.DataSource;
            dt.Rows.Add("", "", "", "D2", "", "False");
            _isDirty = true;
        }

        private void btnRemoveSheet_Click(object sender, EventArgs e)
        {
            if (this.dgvSheets.SelectedRows.Count > 0)
            {
                this.dgvSheets.Rows.RemoveAt(this.dgvSheets.SelectedRows[0].Index);
                _isDirty = true;
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveConfiguration();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            SaveConfiguration();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (_isDirty)
            {
                var result = MessageBox.Show("You have unsaved changes. Do you want to save them before closing?", 
                    "Unsaved Changes", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    SaveConfiguration();
                }
                else if (result == DialogResult.Cancel)
                {
                    return;
                }
            }

            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void SheetConfigManagerForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_isDirty)
            {
                var result = MessageBox.Show("You have unsaved changes. Do you want to save them before closing?", 
                    "Unsaved Changes", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    SaveConfiguration();
                    e.Cancel = false;
                }
                else if (result == DialogResult.Cancel)
                {
                    e.Cancel = true;
                }
            }
        }

        private void SheetConfigManagerForm_Shown(object sender, EventArgs e)
        {
            LoadConfiguration();
            
            // Ensure Sheets tab is selected and visible
            this.tabControl.SelectedTab = this.tabSheets;
            this.tabSheets.Refresh();
            
            System.Diagnostics.Debug.WriteLine("Form shown and configuration loaded");
        }

        // Control declarations
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabSheets;
        private System.Windows.Forms.TabPage tabGeneralSettings;
        private System.Windows.Forms.TabPage tabLoggingSettings;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnApply;

        // Sheets controls
        private System.Windows.Forms.Label lblSheets;
        private System.Windows.Forms.DataGridView dgvSheets;
        private System.Windows.Forms.Button btnAddSheet;
        private System.Windows.Forms.Button btnRemoveSheet;

        // General Settings controls
        private System.Windows.Forms.GroupBox gbGeneralSettings;
        private System.Windows.Forms.CheckBox chkAutoFillCell;
        private System.Windows.Forms.CheckBox chkEnableDebugLog;
        private System.Windows.Forms.NumericUpDown numStartingNumber;
        private System.Windows.Forms.TextBox txtPageBreakColumnName;
        private System.Windows.Forms.TextBox txtEvidenceFontName;
        private System.Windows.Forms.TextBox txtBackButtonFontName;
        private System.Windows.Forms.NumericUpDown numPrintAreaLastRowIdx;
        private System.Windows.Forms.NumericUpDown numColumnWidth;
        private System.Windows.Forms.NumericUpDown numRowHeight;
        private System.Windows.Forms.NumericUpDown numFontSize;
        private System.Windows.Forms.ComboBox cmbPageOrientation;
        private System.Windows.Forms.ComboBox cmbPaperSize;
        private System.Windows.Forms.NumericUpDown numZoom;
        private System.Windows.Forms.CheckBox chkFitToPagesWide;
        private System.Windows.Forms.CheckBox chkFitToPagesTall;
        private System.Windows.Forms.CheckBox chkCenterHorizontally;
        private System.Windows.Forms.NumericUpDown numWindowZoom;
        private System.Windows.Forms.ComboBox cmbViewMode;

        // Logging Settings controls
        private System.Windows.Forms.GroupBox gbLoggingSettings;
        private System.Windows.Forms.TextBox txtLogDirectory;
        private System.Windows.Forms.CheckBox chkEnableFileLogging;
        private System.Windows.Forms.CheckBox chkEnableDebugOutput;
        private System.Windows.Forms.ComboBox cmbLogLevel;
        private System.Windows.Forms.TextBox txtLogFileName;
    }

    // Helper methods for setup
    partial class SheetConfigManagerForm
    {
        private void SetupGeneralSettingsTab()
        {
            this.gbGeneralSettings = new System.Windows.Forms.GroupBox();
            this.chkAutoFillCell = new System.Windows.Forms.CheckBox();
            this.chkEnableDebugLog = new System.Windows.Forms.CheckBox();
            this.numStartingNumber = new System.Windows.Forms.NumericUpDown();
            this.txtPageBreakColumnName = new System.Windows.Forms.TextBox();
            this.txtEvidenceFontName = new System.Windows.Forms.TextBox();
            this.txtBackButtonFontName = new System.Windows.Forms.TextBox();
            this.numPrintAreaLastRowIdx = new System.Windows.Forms.NumericUpDown();
            this.numColumnWidth = new System.Windows.Forms.NumericUpDown();
            this.numRowHeight = new System.Windows.Forms.NumericUpDown();
            this.numFontSize = new System.Windows.Forms.NumericUpDown();
            this.cmbPageOrientation = new System.Windows.Forms.ComboBox();
            this.cmbPaperSize = new System.Windows.Forms.ComboBox();
            this.numZoom = new System.Windows.Forms.NumericUpDown();
            this.chkFitToPagesWide = new System.Windows.Forms.CheckBox();
            this.chkFitToPagesTall = new System.Windows.Forms.CheckBox();
            this.chkCenterHorizontally = new System.Windows.Forms.CheckBox();
            this.numWindowZoom = new System.Windows.Forms.NumericUpDown();
            this.cmbViewMode = new System.Windows.Forms.ComboBox();

            this.gbGeneralSettings.Location = new System.Drawing.Point(6, 6);
            this.gbGeneralSettings.Name = "gbGeneralSettings";
            this.gbGeneralSettings.Size = new System.Drawing.Size(740, 442);
            this.gbGeneralSettings.TabIndex = 0;
            this.gbGeneralSettings.TabStop = false;
            this.gbGeneralSettings.Text = "General Settings";

            // Add basic controls for simplicity
            var panel = new System.Windows.Forms.Panel();
            panel.AutoScroll = true;
            panel.Location = new System.Drawing.Point(6, 19);
            panel.Size = new System.Drawing.Size(728, 417);
            this.gbGeneralSettings.Controls.Add(panel);

            int yPos = 10;
            int labelWidth = 150;
            int controlWidth = 100;
            int height = 20;
            int spacing = 25;

            // Auto Fill Cell
            var lblAutoFill = new System.Windows.Forms.Label();
            lblAutoFill.Text = "Auto Fill Cell:";
            lblAutoFill.Location = new System.Drawing.Point(10, yPos);
            lblAutoFill.Size = new System.Drawing.Size(labelWidth, height);
            lblAutoFill.AutoSize = true;
            panel.Controls.Add(lblAutoFill);

            this.chkAutoFillCell.Text = "";
            this.chkAutoFillCell.Location = new System.Drawing.Point(170, yPos);
            this.chkAutoFillCell.Size = new System.Drawing.Size(controlWidth, height);
            panel.Controls.Add(this.chkAutoFillCell);
            this.chkAutoFillCell.CheckedChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Enable Debug Log
            var lblDebugLog = new System.Windows.Forms.Label();
            lblDebugLog.Text = "Enable Debug Log:";
            lblDebugLog.Location = new System.Drawing.Point(10, yPos);
            lblDebugLog.Size = new System.Drawing.Size(labelWidth, height);
            lblDebugLog.AutoSize = true;
            panel.Controls.Add(lblDebugLog);

            this.chkEnableDebugLog.Text = "";
            this.chkEnableDebugLog.Location = new System.Drawing.Point(170, yPos);
            this.chkEnableDebugLog.Size = new System.Drawing.Size(controlWidth, height);
            panel.Controls.Add(this.chkEnableDebugLog);
            this.chkEnableDebugLog.CheckedChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Starting Number
            var lblStartingNumber = new System.Windows.Forms.Label();
            lblStartingNumber.Text = "Starting Number:";
            lblStartingNumber.Location = new System.Drawing.Point(10, yPos);
            lblStartingNumber.Size = new System.Drawing.Size(labelWidth, height);
            lblStartingNumber.AutoSize = true;
            panel.Controls.Add(lblStartingNumber);

            this.numStartingNumber.Location = new System.Drawing.Point(170, yPos);
            this.numStartingNumber.Size = new System.Drawing.Size(controlWidth, height);
            this.numStartingNumber.Minimum = 1;
            this.numStartingNumber.Maximum = 1000;
            panel.Controls.Add(this.numStartingNumber);
            this.numStartingNumber.ValueChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Page Break Column Name
            var lblPageBreakColumn = new System.Windows.Forms.Label();
            lblPageBreakColumn.Text = "Page Break Column Name:";
            lblPageBreakColumn.Location = new System.Drawing.Point(10, yPos);
            lblPageBreakColumn.Size = new System.Drawing.Size(labelWidth, height);
            lblPageBreakColumn.AutoSize = true;
            panel.Controls.Add(lblPageBreakColumn);

            this.txtPageBreakColumnName.Location = new System.Drawing.Point(170, yPos);
            this.txtPageBreakColumnName.Size = new System.Drawing.Size(controlWidth, height);
            panel.Controls.Add(this.txtPageBreakColumnName);
            this.txtPageBreakColumnName.TextChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Evidence Font Name
            var lblEvidenceFont = new System.Windows.Forms.Label();
            lblEvidenceFont.Text = "Evidence Font Name:";
            lblEvidenceFont.Location = new System.Drawing.Point(10, yPos);
            lblEvidenceFont.Size = new System.Drawing.Size(labelWidth, height);
            lblEvidenceFont.AutoSize = true;
            panel.Controls.Add(lblEvidenceFont);

            this.txtEvidenceFontName.Location = new System.Drawing.Point(170, yPos);
            this.txtEvidenceFontName.Size = new System.Drawing.Size(controlWidth, height);
            panel.Controls.Add(this.txtEvidenceFontName);
            this.txtEvidenceFontName.TextChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Back Button Font Name
            var lblBackButtonFont = new System.Windows.Forms.Label();
            lblBackButtonFont.Text = "Back Button Font Name:";
            lblBackButtonFont.Location = new System.Drawing.Point(10, yPos);
            lblBackButtonFont.Size = new System.Drawing.Size(labelWidth, height);
            lblBackButtonFont.AutoSize = true;
            panel.Controls.Add(lblBackButtonFont);

            this.txtBackButtonFontName.Location = new System.Drawing.Point(170, yPos);
            this.txtBackButtonFontName.Size = new System.Drawing.Size(controlWidth, height);
            panel.Controls.Add(this.txtBackButtonFontName);
            this.txtBackButtonFontName.TextChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Print Area Last Row Index
            var lblPrintAreaRow = new System.Windows.Forms.Label();
            lblPrintAreaRow.Text = "Print Area Last Row Index:";
            lblPrintAreaRow.Location = new System.Drawing.Point(10, yPos);
            lblPrintAreaRow.Size = new System.Drawing.Size(labelWidth, height);
            lblPrintAreaRow.AutoSize = true;
            panel.Controls.Add(lblPrintAreaRow);

            this.numPrintAreaLastRowIdx.Location = new System.Drawing.Point(170, yPos);
            this.numPrintAreaLastRowIdx.Size = new System.Drawing.Size(controlWidth, height);
            this.numPrintAreaLastRowIdx.Minimum = 1;
            this.numPrintAreaLastRowIdx.Maximum = 10000;
            panel.Controls.Add(this.numPrintAreaLastRowIdx);
            this.numPrintAreaLastRowIdx.ValueChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Column Width
            var lblColumnWidth = new System.Windows.Forms.Label();
            lblColumnWidth.Text = "Column Width:";
            lblColumnWidth.Location = new System.Drawing.Point(10, yPos);
            lblColumnWidth.Size = new System.Drawing.Size(labelWidth, height);
            lblColumnWidth.AutoSize = true;
            panel.Controls.Add(lblColumnWidth);

            this.numColumnWidth.Location = new System.Drawing.Point(170, yPos);
            this.numColumnWidth.Size = new System.Drawing.Size(controlWidth, height);
            this.numColumnWidth.DecimalPlaces = 2;
            this.numColumnWidth.Minimum = 0;
            this.numColumnWidth.Maximum = 100;
            panel.Controls.Add(this.numColumnWidth);
            this.numColumnWidth.ValueChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Row Height
            var lblRowHeight = new System.Windows.Forms.Label();
            lblRowHeight.Text = "Row Height:";
            lblRowHeight.Location = new System.Drawing.Point(10, yPos);
            lblRowHeight.Size = new System.Drawing.Size(labelWidth, height);
            lblRowHeight.AutoSize = true;
            panel.Controls.Add(lblRowHeight);

            this.numRowHeight.Location = new System.Drawing.Point(170, yPos);
            this.numRowHeight.Size = new System.Drawing.Size(controlWidth, height);
            this.numRowHeight.DecimalPlaces = 1;
            this.numRowHeight.Minimum = 0;
            this.numRowHeight.Maximum = 100;
            panel.Controls.Add(this.numRowHeight);
            this.numRowHeight.ValueChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Font Size
            var lblFontSize = new System.Windows.Forms.Label();
            lblFontSize.Text = "Font Size:";
            lblFontSize.Location = new System.Drawing.Point(10, yPos);
            lblFontSize.Size = new System.Drawing.Size(labelWidth, height);
            lblFontSize.AutoSize = true;
            panel.Controls.Add(lblFontSize);

            this.numFontSize.Location = new System.Drawing.Point(170, yPos);
            this.numFontSize.Size = new System.Drawing.Size(controlWidth, height);
            this.numFontSize.Minimum = 1;
            this.numFontSize.Maximum = 100;
            panel.Controls.Add(this.numFontSize);
            this.numFontSize.ValueChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Page Orientation
            var lblPageOrientation = new System.Windows.Forms.Label();
            lblPageOrientation.Text = "Page Orientation:";
            lblPageOrientation.Location = new System.Drawing.Point(10, yPos);
            lblPageOrientation.Size = new System.Drawing.Size(labelWidth, height);
            lblPageOrientation.AutoSize = true;
            panel.Controls.Add(lblPageOrientation);

            this.cmbPageOrientation.Location = new System.Drawing.Point(170, yPos);
            this.cmbPageOrientation.Size = new System.Drawing.Size(controlWidth, height);
            this.cmbPageOrientation.Items.AddRange(new object[] { "Portrait", "Landscape" });
            panel.Controls.Add(this.cmbPageOrientation);
            this.cmbPageOrientation.SelectedIndexChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Paper Size
            var lblPaperSize = new System.Windows.Forms.Label();
            lblPaperSize.Text = "Paper Size:";
            lblPaperSize.Location = new System.Drawing.Point(10, yPos);
            lblPaperSize.Size = new System.Drawing.Size(labelWidth, height);
            lblPaperSize.AutoSize = true;
            panel.Controls.Add(lblPaperSize);

            this.cmbPaperSize.Location = new System.Drawing.Point(170, yPos);
            this.cmbPaperSize.Size = new System.Drawing.Size(controlWidth, height);
            this.cmbPaperSize.Items.AddRange(new object[] { "A4", "A3", "Letter", "Legal" });
            panel.Controls.Add(this.cmbPaperSize);
            this.cmbPaperSize.SelectedIndexChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Zoom
            var lblZoom = new System.Windows.Forms.Label();
            lblZoom.Text = "Zoom (%):";
            lblZoom.Location = new System.Drawing.Point(10, yPos);
            lblZoom.Size = new System.Drawing.Size(labelWidth, height);
            lblZoom.AutoSize = true;
            panel.Controls.Add(lblZoom);

            this.numZoom.Location = new System.Drawing.Point(170, yPos);
            this.numZoom.Size = new System.Drawing.Size(controlWidth, height);
            this.numZoom.Minimum = 10;
            this.numZoom.Maximum = 400;
            panel.Controls.Add(this.numZoom);
            this.numZoom.ValueChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Fit to Pages Wide
            var lblFitWide = new System.Windows.Forms.Label();
            lblFitWide.Text = "Fit to Pages Wide:";
            lblFitWide.Location = new System.Drawing.Point(10, yPos);
            lblFitWide.Size = new System.Drawing.Size(labelWidth, height);
            lblFitWide.AutoSize = true;
            panel.Controls.Add(lblFitWide);

            this.chkFitToPagesWide.Text = "";
            this.chkFitToPagesWide.Location = new System.Drawing.Point(170, yPos);
            this.chkFitToPagesWide.Size = new System.Drawing.Size(controlWidth, height);
            panel.Controls.Add(this.chkFitToPagesWide);
            this.chkFitToPagesWide.CheckedChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Fit to Pages Tall
            var lblFitTall = new System.Windows.Forms.Label();
            lblFitTall.Text = "Fit to Pages Tall:";
            lblFitTall.Location = new System.Drawing.Point(10, yPos);
            lblFitTall.Size = new System.Drawing.Size(labelWidth, height);
            lblFitTall.AutoSize = true;
            panel.Controls.Add(lblFitTall);

            this.chkFitToPagesTall.Text = "";
            this.chkFitToPagesTall.Location = new System.Drawing.Point(170, yPos);
            this.chkFitToPagesTall.Size = new System.Drawing.Size(controlWidth, height);
            panel.Controls.Add(this.chkFitToPagesTall);
            this.chkFitToPagesTall.CheckedChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Center Horizontally
            var lblCenterHorizontally = new System.Windows.Forms.Label();
            lblCenterHorizontally.Text = "Center Horizontally:";
            lblCenterHorizontally.Location = new System.Drawing.Point(10, yPos);
            lblCenterHorizontally.Size = new System.Drawing.Size(labelWidth, height);
            lblCenterHorizontally.AutoSize = true;
            panel.Controls.Add(lblCenterHorizontally);

            this.chkCenterHorizontally.Text = "";
            this.chkCenterHorizontally.Location = new System.Drawing.Point(170, yPos);
            this.chkCenterHorizontally.Size = new System.Drawing.Size(controlWidth, height);
            panel.Controls.Add(this.chkCenterHorizontally);
            this.chkCenterHorizontally.CheckedChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Window Zoom
            var lblWindowZoom = new System.Windows.Forms.Label();
            lblWindowZoom.Text = "Window Zoom (%):";
            lblWindowZoom.Location = new System.Drawing.Point(10, yPos);
            lblWindowZoom.Size = new System.Drawing.Size(labelWidth, height);
            lblWindowZoom.AutoSize = true;
            panel.Controls.Add(lblWindowZoom);

            this.numWindowZoom.Location = new System.Drawing.Point(170, yPos);
            this.numWindowZoom.Size = new System.Drawing.Size(controlWidth, height);
            this.numWindowZoom.Minimum = 10;
            this.numWindowZoom.Maximum = 400;
            panel.Controls.Add(this.numWindowZoom);
            this.numWindowZoom.ValueChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // View Mode
            var lblViewMode = new System.Windows.Forms.Label();
            lblViewMode.Text = "View Mode:";
            lblViewMode.Location = new System.Drawing.Point(10, yPos);
            lblViewMode.Size = new System.Drawing.Size(labelWidth, height);
            lblViewMode.AutoSize = true;
            panel.Controls.Add(lblViewMode);

            this.cmbViewMode.Location = new System.Drawing.Point(170, yPos);
            this.cmbViewMode.Size = new System.Drawing.Size(controlWidth, height);
            this.cmbViewMode.Items.AddRange(new object[] { "Normal", "PageBreakPreview", "PageLayout" });
            panel.Controls.Add(this.cmbViewMode);
            this.cmbViewMode.SelectedIndexChanged += new System.EventHandler(this.Control_ValueChanged);

            this.tabGeneralSettings.Controls.Add(this.gbGeneralSettings);
        }

        private void SetupLoggingSettingsTab()
        {
            this.gbLoggingSettings = new System.Windows.Forms.GroupBox();
            this.txtLogDirectory = new System.Windows.Forms.TextBox();
            this.chkEnableFileLogging = new System.Windows.Forms.CheckBox();
            this.chkEnableDebugOutput = new System.Windows.Forms.CheckBox();
            this.cmbLogLevel = new System.Windows.Forms.ComboBox();
            this.txtLogFileName = new System.Windows.Forms.TextBox();

            this.gbLoggingSettings.Location = new System.Drawing.Point(6, 6);
            this.gbLoggingSettings.Name = "gbLoggingSettings";
            this.gbLoggingSettings.Size = new System.Drawing.Size(740, 442);
            this.gbLoggingSettings.TabIndex = 0;
            this.gbLoggingSettings.TabStop = false;
            this.gbLoggingSettings.Text = "Logging Settings";

            int yPos = 25;
            int labelWidth = 120;
            int controlWidth = 200;
            int height = 20;
            int spacing = 30;

            // Log Directory
            var lblLogDirectory = new System.Windows.Forms.Label();
            lblLogDirectory.Text = "Log Directory:";
            lblLogDirectory.Location = new System.Drawing.Point(10, yPos);
            lblLogDirectory.Size = new System.Drawing.Size(labelWidth, height);
            this.gbLoggingSettings.Controls.Add(lblLogDirectory);

            this.txtLogDirectory.Location = new System.Drawing.Point(140, yPos);
            this.txtLogDirectory.Size = new System.Drawing.Size(controlWidth, height);
            this.gbLoggingSettings.Controls.Add(this.txtLogDirectory);
            this.txtLogDirectory.TextChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Enable File Logging
            var lblEnableFileLogging = new System.Windows.Forms.Label();
            lblEnableFileLogging.Text = "Enable File Logging:";
            lblEnableFileLogging.Location = new System.Drawing.Point(10, yPos);
            lblEnableFileLogging.Size = new System.Drawing.Size(labelWidth, height);
            this.gbLoggingSettings.Controls.Add(lblEnableFileLogging);

            this.chkEnableFileLogging.Location = new System.Drawing.Point(140, yPos);
            this.chkEnableFileLogging.Size = new System.Drawing.Size(controlWidth, height);
            this.gbLoggingSettings.Controls.Add(this.chkEnableFileLogging);
            this.chkEnableFileLogging.CheckedChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Enable Debug Output
            var lblEnableDebugOutput = new System.Windows.Forms.Label();
            lblEnableDebugOutput.Text = "Enable Debug Output:";
            lblEnableDebugOutput.Location = new System.Drawing.Point(10, yPos);
            lblEnableDebugOutput.Size = new System.Drawing.Size(labelWidth, height);
            this.gbLoggingSettings.Controls.Add(lblEnableDebugOutput);

            this.chkEnableDebugOutput.Location = new System.Drawing.Point(140, yPos);
            this.chkEnableDebugOutput.Size = new System.Drawing.Size(controlWidth, height);
            this.gbLoggingSettings.Controls.Add(this.chkEnableDebugOutput);
            this.chkEnableDebugOutput.CheckedChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Log Level
            var lblLogLevel = new System.Windows.Forms.Label();
            lblLogLevel.Text = "Log Level:";
            lblLogLevel.Location = new System.Drawing.Point(10, yPos);
            lblLogLevel.Size = new System.Drawing.Size(labelWidth, height);
            this.gbLoggingSettings.Controls.Add(lblLogLevel);

            this.cmbLogLevel.Location = new System.Drawing.Point(140, yPos);
            this.cmbLogLevel.Size = new System.Drawing.Size(controlWidth, height);
            this.cmbLogLevel.Items.AddRange(new object[] { "DEBUG", "INFO", "WARNING", "ERROR" });
            this.gbLoggingSettings.Controls.Add(this.cmbLogLevel);
            this.cmbLogLevel.SelectedIndexChanged += new System.EventHandler(this.Control_ValueChanged);

            yPos += spacing;

            // Log File Name
            var lblLogFileName = new System.Windows.Forms.Label();
            lblLogFileName.Text = "Log File Name:";
            lblLogFileName.Location = new System.Drawing.Point(10, yPos);
            lblLogFileName.Size = new System.Drawing.Size(labelWidth, height);
            this.gbLoggingSettings.Controls.Add(lblLogFileName);

            this.txtLogFileName.Location = new System.Drawing.Point(140, yPos);
            this.txtLogFileName.Size = new System.Drawing.Size(controlWidth, height);
            this.gbLoggingSettings.Controls.Add(this.txtLogFileName);
            this.txtLogFileName.TextChanged += new System.EventHandler(this.Control_ValueChanged);

            this.tabLoggingSettings.Controls.Add(this.gbLoggingSettings);
        }

        private void dgvSheets_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"DataBindingComplete: {this.dgvSheets.Rows.Count} rows displayed");
            
            // Ensure all rows are visible
            if (this.dgvSheets.Rows.Count > 0)
            {
                this.dgvSheets.FirstDisplayedScrollingRowIndex = 0;
            }
        }
    }
}