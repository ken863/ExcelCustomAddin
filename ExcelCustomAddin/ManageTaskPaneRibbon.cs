using Microsoft.Office.Tools.Ribbon;
namespace ExcelCustomAddin
{
    public partial class ManageTaskPaneRibbon
    {
        /// <summary>
        /// btnTranslate_Click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTranslate_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.myCustomTaskPane.Visible = !Globals.ThisAddIn.myCustomTaskPane.Visible;
        }
    }
}
