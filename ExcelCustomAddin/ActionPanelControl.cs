using System;
using System.Windows.Forms;

namespace ExcelCustomAddin
{
    public partial class ActionPanelControl : UserControl
    {
        /// <summary>
        /// ActionPanelControl
        /// </summary>
        public ActionPanelControl()
        {
            InitializeComponent();
            txtApiKey.Text = Properties.Settings.Default.API_KEY;
            txtModel.Text = Properties.Settings.Default.MODEL;
        }

        public event EventHandler TranslateSheetEvent;
        public event EventHandler TranslateSelectedEvent;

        /// <summary>
        /// btnTranslate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTranslateSelectedText_Click(object sender, EventArgs e)
        {
            if (this.TranslateSelectedEvent != null)
                this.TranslateSelectedEvent(this, e);
        }

        /// <summary>
        /// btnSheetTranslate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSheetTranslate_Click(object sender, EventArgs e)
        {
            if (this.TranslateSheetEvent != null)
                this.TranslateSheetEvent(this, e);
        }

        /// <summary>
        /// txtApiKey_TextChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtApiKey_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.API_KEY = txtApiKey.Text.Trim();
            Properties.Settings.Default.Save();
        }

        /// <summary>
        /// txtModel_TextChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtModel_TextChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.MODEL = txtModel.Text.Trim();
            Properties.Settings.Default.Save();
        }
    }
}
