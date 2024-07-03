using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelCustomAddin
{
    public partial class ActionPanelControl : UserControl
    {
        public ActionPanelControl()
        {
            InitializeComponent();
        }

        public event EventHandler ListOfSheet_SelectedIndexChanged;
        public event EventHandler TranslateClick;

        /// <summary>
        /// listOfSheet_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listOfSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.ListOfSheet_SelectedIndexChanged != null)
                this.ListOfSheet_SelectedIndexChanged(this, e);
        }

        /// <summary>
        /// btnTranslate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTranslate_Click(object sender, EventArgs e)
        {
            if (this.TranslateClick != null)
                this.TranslateClick(this, e);
        }
    }
}
