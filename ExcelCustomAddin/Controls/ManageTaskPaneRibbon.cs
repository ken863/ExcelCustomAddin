using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows;

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
            try
            {
                Globals.ThisAddIn.myCustomTaskPane.Visible = !Globals.ThisAddIn.myCustomTaskPane.Visible;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// btnSheetConfigManager_Click
        /// Mở Sheet Configuration Manager form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSheetConfigManager_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var configForm = new Controls.SheetConfigManagerForm();
                configForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening Sheet Configuration Manager: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
