namespace ExcelCustomAddin
{
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Text;

    public partial class ThisAddIn
    {
        private ActionPanelControl _actionPanel { get; set; }

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

        private bool IsSheetActivating { get; set; } = false;

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
            Globals.ThisAddIn.Application.SheetActivate += Application_SheetActivate;

            this.CreateActionsPane(this.Application.ActiveWorkbook);
        }

        private void Application_SheetActivate(object Sh)
        {
            this.IsSheetActivating = true;
            _actionPanel.listofSheet.DataSource = this.GetListOfSheet();
            _actionPanel.listofSheet.SelectedIndex = FindIndexOfSelectedSheet();
            this.IsSheetActivating = false;

        }

        private List<string> GetListOfSheet()
        {
            return (from Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets select sheet.Name).ToList();
        }

        private int FindIndexOfSelectedSheet()
        {
            return _actionPanel.listofSheet.Items.Cast<string>().ToList().FindIndex(item => item == Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Name);
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
                _actionPanel = (ActionPanelControl)myCustomTaskPane?.Control;
                _actionPanel.listofSheet.DataSource = this.GetListOfSheet();
                _actionPanel.listofSheet.SelectedIndexChanged -= this.ListOfSheet_SelectionChanged;
                _actionPanel.listofSheet.SelectedIndexChanged += this.ListOfSheet_SelectionChanged;
                _actionPanel.TranslateSheetEvent += this.TranslateSheetAsync;
            }
        }

        private void ListOfSheet_SelectionChanged(object sender, EventArgs e)
        {
            if (this.IsSheetActivating)
            {
                return;
            }

            this.SetActiveSheet();
        }

        private async void TranslateSheetAsync(object sender, EventArgs e)
        {
            var sheetValue = this.GetSheetValues();

            if (sheetValue != null)
            {
                var apiKey = "sk-proj-KHBw6jj2cKclN3xmD5olT3BlbkFJekvhNIP9ykw0F1xIScCD";
                var chatGPTClient = new ChatGPTClient(apiKey);
                var response = await chatGPTClient.CallChatGPTAsync(sheetValue);

                foreach (string line in response.Split('\n'))
                {
                    var arr = line.Split('|');
                    if (arr.Length > 1)
                    {
                        var cellAddress = arr[0];
                        var cellValue = arr[1];

                        Range range = Application.ActiveSheet.Range[cellAddress];
                        range.Value2 = cellValue;
                    }
                }
            }
        }

        private void SetActiveSheet()
        {
            var selectedSheetName = _actionPanel.listofSheet.SelectedValue?.ToString();
            if (!string.IsNullOrEmpty(selectedSheetName))
            {
                // Sử dụng LINQ để tìm worksheet theo tên
                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets
                    .Cast<Worksheet>()
                    .FirstOrDefault(ws => ws.Name == selectedSheetName);

                if (sheet != null)
                {
                    // Đặt worksheet này là active sheet
                    sheet.Activate();
                }
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
            StringBuilder sb = new StringBuilder();

            var selectedRange = Globals.ThisAddIn.Application.Selection;

            // Kiểm tra xem có bất kỳ range nào đang được chọn không
            if (selectedRange != null)
            {
                BackgroundWorker backgroundWorker = new BackgroundWorker();

                // Đăng ký sự kiện DoWork và RunWorkerCompleted
                backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);

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

        private string GetSheetValues()
        {
            Worksheet worksheet = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            // Lấy toàn bộ các ô trong worksheet
            Range usedRange = worksheet.UsedRange.Cells;

            // Dùng LINQ để lấy địa chỉ của các ô có giá trị
            var nonEmptyCellAddresses = (from Range cell in usedRange
                                         where cell.Value2 != null
                                         select cell.Address[false, false] + "| " + cell.Value2?.ToString().Trim()).ToList();

            var result = string.Join("\n", nonEmptyCellAddresses);

            return result;
        }
    }
}