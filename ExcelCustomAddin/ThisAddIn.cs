namespace ExcelCustomAddin
{
    using Microsoft.Office.Core;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using System;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Forms;
    using Excel = Microsoft.Office.Interop.Excel;

    public partial class ThisAddIn
    {
        private ActionPanelControl _actionPanel { get; set; }

        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane { get; set; }

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
            var apiKey = "sk-proj-KHBw6jj2cKclN3xmD5olT3BlbkFJekvhNIP9ykw0F1xIScCD";
            chatGPTClient = new ChatGPTClient(apiKey);

            if (this.Application.ActiveWorkbook != null)
            {
                InitWorbBook(this.Application.ActiveWorkbook);

                // Register Events
                this.Application.ActiveWorkbook.NewSheet += new WorkbookEvents_NewSheetEventHandler(WorkBook_NewSheet);
                this.Application.ActiveWorkbook.SheetBeforeDelete += new WorkbookEvents_SheetBeforeDeleteEventHandler(WorkBook_SheetBeforeDelete);
                this.Application.ActiveWorkbook.SheetActivate += new WorkbookEvents_SheetActivateEventHandler(WorkBook_SheetActivate);
                this.Application.ActiveWorkbook.SheetSelectionChange += new WorkbookEvents_SheetSelectionChangeEventHandler(WorkBook_SheetSelectionChange);

                // Get Active ActionsPanel
                _actionPanel.ListOfSheet_SelectedIndexChanged += new EventHandler(ListOfSheet_SelectedIndexChanged);
                _actionPanel.VisibleChanged += new EventHandler(ListOfSheet_VisibleChangedd);
                _actionPanel.TranslateClick += new EventHandler(TranslateClickAsync);

                // Make list of worksheet on ActionsPane
                CreateWorkSheetList();
            }

            // Register Hanle Events
            ((AppEvents_Event)Application).NewWorkbook += new AppEvents_NewWorkbookEventHandler(Application_NewWorkbook);
            this.Application.WorkbookOpen += new AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            this.Application.WorkbookActivate += new AppEvents_WorkbookActivateEventHandler(Application_WorkbookActive);
            this.Application.WorkbookBeforeClose += new AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookBeforeClose);
        }

        private void ListOfSheet_VisibleChangedd(object sender, EventArgs e)
        {
            if (_actionPanel == null)
            {
                return;
            }

            Globals.Ribbons.ManageTaskPaneRibbon.toggleButton1.Checked = myCustomTaskPane.Visible;
        }

        /// <summary>
        /// Application_NewWorkbook
        /// </summary>
        /// <param name="Wb"></param>
        void Application_NewWorkbook(Workbook Wb)
        {
            RegisterEvents(Wb);

            InitWorbBook(Wb);
        }

        /// <summary>
        /// Application_WorkbookActivate
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookOpen(Workbook Wb)
        {
            RegisterEvents(Wb);

            InitWorbBook(Wb);
        }

        /// <summary>
        /// Application_WorkbookBeforeClose
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookBeforeClose(Workbook Wb, ref bool isCancel)
        {
            if (isCancel || Wb == null)
            {
                return;
            }

            // UnRegister Events
            Wb.NewSheet -= new WorkbookEvents_NewSheetEventHandler(WorkBook_NewSheet);
            Wb.SheetBeforeDelete -= new WorkbookEvents_SheetBeforeDeleteEventHandler(WorkBook_SheetBeforeDelete);
            Wb.SheetActivate -= new WorkbookEvents_SheetActivateEventHandler(WorkBook_SheetActivate);
        }

        /// <summary>
        /// Application_WorkbookActivate
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookActive(Workbook Wb)
        {
            if (Wb == null)
            {
                return;
            }

            RegisterEvents(Wb);

            // Get Active ActionsPanel
            _actionPanel = (ActionPanelControl)this.CustomTaskPanes.Where(_ => _.Title == Wb.Name).FirstOrDefault().Control;
            _actionPanel.ListOfSheet_SelectedIndexChanged -= new EventHandler(ListOfSheet_SelectedIndexChanged);
            _actionPanel.ListOfSheet_SelectedIndexChanged += new EventHandler(ListOfSheet_SelectedIndexChanged);
            _actionPanel.TranslateClick -= new EventHandler(TranslateClickAsync);
            _actionPanel.TranslateClick += new EventHandler(TranslateClickAsync);
            _actionPanel.VisibleChanged -= new EventHandler(ListOfSheet_VisibleChangedd);
            _actionPanel.VisibleChanged += new EventHandler(ListOfSheet_VisibleChangedd);

            myCustomTaskPane = TaskPaneManager.GetTaskPane(Wb.Name, () => new ActionPanelControl());
        }

        private void RegisterEvents(Workbook Wb)
        {
            if (Wb == null)
            {
                return;
            }

            Wb.NewSheet -= new WorkbookEvents_NewSheetEventHandler(WorkBook_NewSheet);
            Wb.SheetBeforeDelete -= new WorkbookEvents_SheetBeforeDeleteEventHandler(WorkBook_SheetBeforeDelete);
            Wb.SheetActivate -= new WorkbookEvents_SheetActivateEventHandler(WorkBook_SheetActivate);
            Wb.SheetSelectionChange -= new WorkbookEvents_SheetSelectionChangeEventHandler(WorkBook_SheetSelectionChange);

            // Register Events
            Wb.NewSheet += new WorkbookEvents_NewSheetEventHandler(WorkBook_NewSheet);
            Wb.SheetBeforeDelete += new WorkbookEvents_SheetBeforeDeleteEventHandler(WorkBook_SheetBeforeDelete);
            Wb.SheetActivate += new WorkbookEvents_SheetActivateEventHandler(WorkBook_SheetActivate);
            Wb.SheetSelectionChange += new WorkbookEvents_SheetSelectionChangeEventHandler(WorkBook_SheetSelectionChange);
        }

        /// <summary>
        /// WorkBook_NewSheet
        /// </summary>
        /// <param name="sheet"></param>
        private void WorkBook_SheetActivate(object sheet)
        {
            if (sheet == null)
            {
                return;
            }

            CreateWorkSheetList();

            Worksheet worksheet = (Excel.Worksheet)sheet;

            for (int i = 0; i < _actionPanel.listOfSheet.Items.Count; i++)
            {
                if ((string)_actionPanel.listOfSheet.Items[i] == worksheet.Name)
                {
                    _actionPanel.listOfSheet.SelectedIndex = i;
                    break;
                }
            }
        }

        /// <summary>
        /// WorkBook_SheetSelectionChange
        /// </summary>
        /// <param name="sheet"></param>
        private void WorkBook_SheetSelectionChange(object Sh, Range Target)
        {
            if (Target == null)
            {
                return;
            }

            GetSelectedText();
        }

        /// <summary>
        /// WorkBook_NewSheet
        /// </summary>
        /// <param name="sheet"></param>
        private void WorkBook_NewSheet(Object sheet)
        {
            CreateWorkSheetList();
        }

        /// <summary>
        /// WorkBook_SheetBeforeDelete
        /// </summary>
        /// <param name="sheet"></param>
        private void WorkBook_SheetBeforeDelete(Object sheet)
        {
            CreateWorkSheetList();
        }

        /// <summary>
        /// CreateWorkSheetList
        /// </summary>
        private void CreateWorkSheetList()
        {
            if (_actionPanel == null || this.Application.ActiveWorkbook == null)
            {
                return;
            }

            // Xoá toàn bộ sheet
            _actionPanel.listOfSheet.Items.Clear();

            // Lấy đối tượng workbook hiện tại
            Workbook workbook = this.Application.ActiveWorkbook;

            if (workbook != null)
            {
                // Lấy collection các sheet trong workbook
                Sheets sheets = workbook.Sheets;

                // Duyệt qua tất cả các sheet và lấy tên của chúng
                foreach (Excel.Worksheet sheet in sheets)
                {
                    _actionPanel.listOfSheet.Items.Add(sheet.Name);
                }
            }
        }

        /// <summary>
        /// AddActionPaneToWorkbook
        /// </summary>
        /// <param name="workbook"></param>
        private void InitWorbBook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            myCustomTaskPane = TaskPaneManager.GetTaskPane(workbook.Name, () => new ActionPanelControl());
            myCustomTaskPane.Visible = false;

            // Get Active ActionsPanel
            _actionPanel = (ActionPanelControl)this.CustomTaskPanes.Where(_ => _.Title == workbook.Name).FirstOrDefault().Control;

            CreateWorkSheetList();
        }

        /// <summary>
        /// ListOfSheet_SelectedIndexChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListOfSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (_actionPanel.listOfSheet.SelectedItem == null)
                {
                    return;
                }

                Worksheet sheet = null;
                string sheetName = _actionPanel.listOfSheet.SelectedItem.ToString();

                if (!string.IsNullOrEmpty(sheetName.Trim()))
                {
                    Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                    foreach (Excel.Worksheet ws in workbook.Sheets)
                    {
                        if (ws.Name == sheetName)
                        {
                            sheet = ws;
                            break;
                        }
                    }

                    if (sheet != null)
                    {
                        // Activate the sheet
                        sheet.Activate();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// TranslateClick
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void TranslateClickAsync(object sender, EventArgs e)
        {
            try
            {
                var response = await chatGPTClient.CallChatGPTAsync(_actionPanel.txtSourceText.Text);
                this.UpdateText(response);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Hàm cập nhật text trên txtDesText
        private void UpdateText(string text)
        {
            if (_actionPanel.txtDesText.InvokeRequired)
            {
                // Nếu cần phải gọi Invoke, sử dụng phương thức này để gọi hàm từ thread khác
                _actionPanel.txtDesText.Invoke(new Action<string>(UpdateText), text);
            }
            else
            {
                // Nếu không cần phải gọi Invoke, cập nhật trực tiếp
                _actionPanel.txtDesText.Text = text;
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