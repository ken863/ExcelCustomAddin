namespace ExcelCustomAddin
{
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using System;
    using System.Linq;
    using System.Text;
    using System.Windows.Forms;
    using Excel = Microsoft.Office.Interop.Excel;

    public partial class ThisAddIn
    {
        private ActionPanelControl _actionPanel
        {
            get
            {
                var currentActionPane = (ActionPanelControl)myCustomTaskPane?.Control;
                return currentActionPane;
            }
        }

        /// <summary>
        /// myCustomTaskPane
        /// </summary>
        public CustomTaskPane myCustomTaskPane { get; set; }

        /// <summary>
        /// chatGPTClient
        /// </summary>
        ChatGPTClient chatGPTClient { get; set; }

        /// <summary>
        /// InternalStartup
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
        }

        /// <summary>
        /// ThisAddIn_Startup
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            // Register Hanle Events
            ((AppEvents_Event)Globals.ThisAddIn.Application).NewWorkbook += Application_NewWorkbook;
            Globals.ThisAddIn.Application.WorkbookOpen += Application_WorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookActivate += Application_WorkbookActive;
            Globals.ThisAddIn.Application.SheetSelectionChange += Application_SheetSelectionChange;
        }

        /// <summary>
        /// Application_NewWorkbook
        /// </summary>
        /// <param name="Wb"></param>
        void Application_NewWorkbook(Workbook Wb)
        {
            this.CreateActionsPane(Wb);
        }

        /// <summary>
        /// Application_WorkbookActivate
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookOpen(Workbook Wb)
        {
            this.CreateActionsPane(Wb);
        }

        /// <summary>
        /// Application_WorkbookActivate
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookActive(Workbook Wb)
        {
            this.CreateActionsPane(Wb);
        }

        private void CreateActionsPane(Workbook Wb)
        {
            if (Wb != null)
            {
                // Get Active ActionsPanel
                myCustomTaskPane = TaskPaneManager.GetTaskPane(Wb.Name, "TRANSLATE TOOL", () => new ActionPanelControl());
            }
        }

        /// <summary>
        /// Application_SheetSelectionChange
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        private void Application_SheetSelectionChange(object sh, Range target)
        {
            this.GetSelectedText();
        }

        /// <summary>
        /// ActionsPane_ClickTranslate
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ActionsPane_ClickTranslate(object sender, EventArgs e)
        {
            try
            {
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void GetSelectedText()
        {
            // Giả sử rằng 'application' là biến đại diện cho đối tượng ứng dụng Excel hiện tại (Application)
            Excel.Application application = Globals.ThisAddIn.Application;

            StringBuilder sb = new StringBuilder();

            // Kiểm tra xem có bất kỳ range nào đang được chọn không
            if (application.Selection != null)
            {
                // Lấy ra đối tượng Selection
                Range selectedRange = application.Selection as Range;

                if (selectedRange != null)
                {
                    // Lặp qua từng hàng trong vùng
                    foreach (Range row in selectedRange.Rows)
                    {
                        string rowData = "";
                        // Lấy số lượng cột trong hàng
                        int columnCount = row.Columns.Count;

                        for (int i = 1; i <= columnCount; i++)
                        {
                            // Lấy giá trị của mỗi ô
                            Range cell = row.Cells[1, i];
                            string cellValue = cell.Value?.ToString() ?? "";

                            rowData += cellValue.Trim() + "\t";
                        }

                        sb.AppendLine(rowData);
                    }

                    _actionPanel.txtSourceText.Text = sb.ToString().Trim();
                }
            }
        }
    }
}