using System;
using System.Windows.Forms;

namespace ExcelCustomAddin
{
    public partial class ActionPanelControl : UserControl
    {

        public ActionPanelControl()
        {
            InitializeComponent();
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
    }
}
