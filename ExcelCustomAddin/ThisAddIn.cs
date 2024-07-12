namespace ExcelCustomAddin
{
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Text;
    using System.Windows;
    using System.Windows.Threading;

    public partial class ThisAddIn
    {
        /// <summary>
        /// ActionPanelControl
        /// </summary>
        private ActionPanelControl _actionPanel { get; set; }

        /// <summary>
        /// CustomTaskPane
        /// </summary>
        public CustomTaskPane myCustomTaskPane { get; set; }

        /// <summary>
        /// BackgroundWorker
        /// </summary>
        private BackgroundWorker backgroundWorker;

        /// <summary>
        /// Dispatcher
        /// </summary>
        private Dispatcher _dispatcher;

        /// <summary>
        /// Kiểm tra sheet đang trong quá trình kích hoạt 
        /// để ngăn chặn sự kiện update danh sách sheet khi SelectedIndexChanged và SheetActivate
        /// </summary>
        private bool IsSheetActivating { get; set; } = false;

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
            // Tạo Dispatcher từ thread chính của ứng dụng
            _dispatcher = Dispatcher.CurrentDispatcher;

            // Register Hanle Events
            ((AppEvents_Event)Globals.ThisAddIn.Application).NewWorkbook += Application_NewWorkbook;
            Globals.ThisAddIn.Application.WorkbookOpen += Application_WorkbookOpen;
            Globals.ThisAddIn.Application.WorkbookActivate += Application_WorkbookActive;
            Globals.ThisAddIn.Application.SheetSelectionChange += Application_SheetSelectionChange;
            Globals.ThisAddIn.Application.SheetActivate += Application_SheetActivate;

            // Tạo ActionPane
            this.CreateActionsPane(this.Application.ActiveWorkbook);
        }

        #region "Quản lý ActionPane"
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
                _actionPanel.TranslateSelectedEvent += this.TranslateSelectedEvent;
            }
        }
        #endregion

        #region "Chức năng tạo danh sách sheet"
        /// <summary>
        /// ListOfSheet_SelectionChanged
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ListOfSheet_SelectionChanged(object sender, EventArgs e)
        {
            if (this.IsSheetActivating)
            {
                return;
            }

            this.SetActiveSheet();
        }

        /// <summary>
        /// SetActiveSheet
        /// </summary>
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
        /// Application_SheetActivate
        /// </summary>
        /// <param name="Sh"></param>
        private void Application_SheetActivate(object Sh)
        {
            this.IsSheetActivating = true;
            _actionPanel.listofSheet.DataSource = this.GetListOfSheet();
            _actionPanel.listofSheet.SelectedIndex = FindIndexOfSelectedSheet();
            this.IsSheetActivating = false;
        }

        /// <summary>
        /// GetListOfSheet
        /// </summary>
        /// <returns></returns>
        private List<string> GetListOfSheet()
        {
            return (from Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets select sheet.Name).ToList();
        }

        /// <summary>
        /// FindIndexOfSelectedSheet
        /// </summary>
        /// <returns></returns>
        private int FindIndexOfSelectedSheet()
        {
            return _actionPanel.listofSheet.Items.Cast<string>().ToList()
                .FindIndex(item => item == Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Name);
        }
        #endregion

        #region "Chức năng chọn cell"
        /// <summary>
        /// Application_SheetSelectionChange
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="target"></param>
        private void Application_SheetSelectionChange(object sh, Range target)
        {
            if (myCustomTaskPane.Visible)
            {
                StringBuilder sb = new StringBuilder();

                var selectedRange = Globals.ThisAddIn.Application.Selection;

                // Kiểm tra xem có bất kỳ range nào đang được chọn không
                if (selectedRange != null)
                {
                    // Đăng ký sự kiện DoWork và RunWorkerCompleted
                    backgroundWorker = new BackgroundWorker();
                    backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoSelectText);

                    // Bắt đầu BackgroundWorker
                    backgroundWorker.RunWorkerAsync(selectedRange);
                }
            }
        }

        /// <summary>
        /// BackgroundWorker_DoSelectText
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BackgroundWorker_DoSelectText(object sender, DoWorkEventArgs e)
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

                string result = string.Join(Environment.NewLine, rangeValues);

                if (!string.IsNullOrEmpty(result))
                {
                    _dispatcher.Invoke(new System.Action(() =>
                    {
                        // Nếu không cần phải gọi Invoke, cập nhật trực tiếp
                        _actionPanel.txtSourceText.Text = result.Trim();
                    }));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        #endregion

        #region "Chức năng dịch sheet"
        /// <summary>
        /// TranslateSheetAsync
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TranslateSheetAsync(object sender, EventArgs e)
        {
            // Vô hiệu các Control
            EnableControl(false);

            // Lấy toàn bộ các ô đang sử dụng trong worksheet
            Worksheet worksheet = (Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Range usedRange = worksheet.UsedRange.Cells;

            // Đăng ký sự kiện DoWork và RunWorkerCompleted
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoTranslateSheet);

            // Bắt đầu BackgroundWorker
            backgroundWorker.RunWorkerAsync(usedRange);
        }

        /// <summary>
        /// BackgroundWorker_DoTranslateSheetAsync
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BackgroundWorker_DoTranslateSheet(object sender, DoWorkEventArgs e)
        {
            var selectedRange = (Range)e.Argument;
            var sheetValue = this.GetSheetValues(selectedRange);

            if (sheetValue != null)
            {
                var chatGPTClient = new ChatGPTClient();
                var response = await chatGPTClient.CallChatGPTAsync(sheetValue);

                foreach (string line in response.Split('\n'))
                {
                    var arr = line.Split('|');
                    if (arr.Length > 1)
                    {
                        var cellAddress = arr[0];
                        var cellValue = arr[1];

                        _dispatcher.Invoke(new System.Action(() =>
                        {
                            Range range = Application.ActiveSheet.Range[cellAddress];
                            range.Value2 = cellValue;
                        }));
                    }
                }
            }

            _dispatcher.Invoke(new System.Action(() =>
            {
                // Kích hoạt các Control
                EnableControl(true);
            }));
        }

        /// <summary>
        /// GetSheetValues
        /// </summary>
        /// <returns></returns>
        private string GetSheetValues(Range usedRange)
        {
            // Dùng LINQ để lấy địa chỉ của các ô có giá trị
            var nonEmptyCellAddresses = (from Range cell in usedRange
                                         where cell.Value2 != null
                                         select cell.Address[false, false] + "| " + cell.Value2?.ToString().Trim()).ToList();

            var result = string.Join(Environment.NewLine, nonEmptyCellAddresses);

            return result;
        }
        #endregion

        #region "Chức năng dịch text đã chọn"
        /// <summary>
        /// TranslateSelectedEventAsync
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TranslateSelectedEvent(object sender, EventArgs e)
        {
            // Vô hiệu các Control
            EnableControl(false);

            // Lấy text
            var selectedValue = _actionPanel.txtSourceText.Text.Trim();

            // Đăng ký sự kiện DoWork và RunWorkerCompleted
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoTranslateSelectedText);

            // Bắt đầu BackgroundWorker
            backgroundWorker.RunWorkerAsync(selectedValue);
        }

        /// <summary>
        /// BackgroundWorker_DoTranslateSelectedTextAsync
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BackgroundWorker_DoTranslateSelectedText(object sender, DoWorkEventArgs e)
        {
            try
            {
                var selectedValue = (string)e.Argument;

                if (!string.IsNullOrEmpty(selectedValue))
                {
                    var chatGPTClient = new ChatGPTClient();
                    var response = await chatGPTClient.CallChatGPTAsync(selectedValue);

                    _dispatcher.Invoke(new System.Action(() =>
                    {
                        // Setting giá trị sau khi dịch
                        _actionPanel.txtDesText.Text = response.Replace("\n", Environment.NewLine);
                    }));
                }

                _dispatcher.Invoke(new System.Action(() =>
                {
                    // Kích hoạt các Control
                    EnableControl(true);
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        /// <summary>
        /// EnableControl
        /// </summary>
        /// <param name="enable"></param>
        private void EnableControl(bool enable)
        {
            _actionPanel.txtSourceText.Enabled = enable;
            _actionPanel.txtDesText.Enabled = enable;
            _actionPanel.btnSheetTranslate.Enabled = enable;
            _actionPanel.btnTranslateSelectedText.Enabled = enable;
            _actionPanel.progressBar.Visible = !enable;
        }
    }
}