namespace ExcelCustomAddin
{
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using System;
    using System.ComponentModel;
    using System.Linq;
    using System.Text;
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

        BackgroundWorker backgroundWorker = new BackgroundWorker();

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

            this.CreateActionsPane(this.Application.ActiveWorkbook);
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
            if (myCustomTaskPane.Visible)
            {
                this.GetSelectedText();
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
                BackgroundWorker backgroundWorker = new BackgroundWorker();

                // Đăng ký sự kiện DoWork và RunWorkerCompleted
                backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);

                Range selectedRange = application.Selection as Range;

                // Bắt đầu BackgroundWorker
                backgroundWorker.RunWorkerAsync(selectedRange);
            }
        }

        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                var selectedRange = (Range)e.Argument;

                // Kiểm tra xem toàn bộ hàng đang được chọn
                if (selectedRange.Rows.Count > 1 && selectedRange.Rows.Count == this.Application.Rows.Count)
                {
                    // Một hàng đang được chọn
                    return;
                }

                // Kiểm tra xem toàn bộ hàng đang được chọn
                if (selectedRange.Columns.Count > 1 && selectedRange.Columns.Count == this.Application.Columns.Count)
                {
                    // Một hàng đang được chọn
                    return;
                }

                var rangeValues = selectedRange.Cells.Cast<Range>().Select(cell => cell.Value2?.ToString().Trim())
                             .Where(value => !string.IsNullOrEmpty(value));

                string result = string.Join("\n", rangeValues);

                if (!string.IsNullOrEmpty(result))
                {
                    this.UpdateText(result.Trim());
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// UpdateText
        /// </summary>
        /// <param name="text"></param>
        private void UpdateText(string text)
        {
            if (_actionPanel.txtSourceText.InvokeRequired)
            {
                // Nếu cần phải gọi Invoke, sử dụng phương thức này để gọi hàm từ thread khác
                _actionPanel.txtSourceText.Invoke(new Action<string>(UpdateText), text);
            }
            else
            {
                // Nếu không cần phải gọi Invoke, cập nhật trực tiếp
                _actionPanel.txtSourceText.Text = text;
            }
        }
    }
}