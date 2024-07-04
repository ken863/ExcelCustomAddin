using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCustomAddin
{
    public partial class ManageTaskPaneRibbon
    {
        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.myCustomTaskPane == null)
            {
                return;

            }

            Globals.ThisAddIn.myCustomTaskPane.Visible = ((RibbonToggleButton)sender).Checked;
        }
    }
}
