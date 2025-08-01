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
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        /// <summary>
        /// ThisAddIn_Shutdown - Cleanup events
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // Hủy đăng ký các application events
                if (Globals.ThisAddIn.Application != null)
                {
                    ((AppEvents_Event)Globals.ThisAddIn.Application).NewWorkbook -= Application_NewWorkbook;
                    Globals.ThisAddIn.Application.WorkbookOpen -= Application_WorkbookOpen;
                    Globals.ThisAddIn.Application.WorkbookActivate -= Application_WorkbookActive;
                    Globals.ThisAddIn.Application.SheetActivate -= Application_SheetActivate;
                }

                // Hủy đăng ký action panel events
                if (_actionPanel != null)
                {
                    _actionPanel.CreateEvidenceEvent -= this.CreateEvidence;
                    _actionPanel.FormatDocumentEvent -= this.FormatDocument;
                    _actionPanel.ChangeSheetNameEvent -= this.ChangeSheetName;
                    _actionPanel.listofSheet.SelectedIndexChanged -= this.ListOfSheet_SelectionChanged;
                }
            }
            catch (Exception ex)
            {
                // Log error if needed, but don't show MessageBox during shutdown
                System.Diagnostics.Debug.WriteLine($"Error during shutdown: {ex.Message}");
            }
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
                myCustomTaskPane = TaskPaneManager.GetTaskPane(Wb.Name, "WORKSHEET TOOLS", () => new ActionPanelControl());
                _actionPanel = (ActionPanelControl)myCustomTaskPane?.Control;

                if (_actionPanel != null)
                {
                    // Hủy đăng ký các event cũ trước khi đăng ký mới để tránh đăng ký trùng lặp
                    _actionPanel.CreateEvidenceEvent -= this.CreateEvidence;
                    _actionPanel.FormatDocumentEvent -= this.FormatDocument;
                    _actionPanel.ChangeSheetNameEvent -= this.ChangeSheetName;
                    _actionPanel.listofSheet.SelectedIndexChanged -= this.ListOfSheet_SelectionChanged;

                    // Đăng ký các event mới
                    _actionPanel.CreateEvidenceEvent += this.CreateEvidence;
                    _actionPanel.FormatDocumentEvent += this.FormatDocument;
                    _actionPanel.ChangeSheetNameEvent += this.ChangeSheetName;
                    _actionPanel.listofSheet.SelectedIndexChanged += this.ListOfSheet_SelectionChanged;

                    // Cập nhật danh sách sheet
                    _actionPanel.listofSheet.DataSource = this.GetListOfSheet();
                }
            }
        }

        /// <summary>
        /// ChangeSheetName
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChangeSheetName(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var activeWorkbook = app.ActiveWorkbook;

                // Kiểm tra workbook
                if (activeWorkbook == null)
                {
                    MessageBox.Show("Không có workbook nào đang mở. Vui lòng mở một workbook và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Lấy tên sheet từ listbox thay vì active sheet
                var selectedItem = _actionPanel.listofSheet.SelectedItem as SheetInfo;
                if (selectedItem == null || string.IsNullOrEmpty(selectedItem.Name))
                {
                    MessageBox.Show("Vui lòng chọn một sheet từ danh sách để đổi tên.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string selectedSheetName = selectedItem.Name;

                // Tìm worksheet theo tên được chọn
                Worksheet selectedSheet = null;
                foreach (Worksheet ws in activeWorkbook.Worksheets)
                {
                    if (ws.Name == selectedSheetName)
                    {
                        selectedSheet = ws;
                        break;
                    }
                }

                if (selectedSheet == null)
                {
                    MessageBox.Show($"Không tìm thấy sheet có tên '{selectedSheetName}'.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Kiểm tra sheet có bị bảo vệ không
                if (selectedSheet.ProtectContents || selectedSheet.ProtectDrawingObjects || selectedSheet.ProtectScenarios)
                {
                    MessageBox.Show($"Sheet '{selectedSheetName}' đang được bảo vệ. Vui lòng bỏ bảo vệ sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string oldSheetName = selectedSheet.Name;

                // Hiển thị dialog để nhập tên mới
                string newSheetName = "";
                if (System.Windows.MessageBox.Show($"Bạn có muốn đổi tên sheet '{oldSheetName}' không?",
                    "Xác nhận", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    // Sử dụng Application.InputBox của Excel thay vì Microsoft.VisualBasic
                    object result = app.InputBox(
                        $"Nhập tên mới cho sheet '{oldSheetName}':",
                        "Đổi tên Sheet",
                        oldSheetName,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2 // Type 2 = text
                    );

                    // Kiểm tra user có hủy không
                    // Check if the user canceled the input box
                    if (result == null || result.ToString() == string.Empty)
                    {
                        return; // User canceled or provided an empty input
                    }

                    newSheetName = result.ToString().Trim();
                }

                // Kiểm tra user có nhập tên mới không
                if (string.IsNullOrWhiteSpace(newSheetName) || newSheetName == oldSheetName)
                {
                    return; // User hủy hoặc không thay đổi
                }

                // Kiểm tra tên sheet mới có hợp lệ không
                if (newSheetName.Length > 31)
                {
                    MessageBox.Show("Tên sheet không được vượt quá 31 ký tự.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Kiểm tra ký tự không hợp lệ
                char[] invalidChars = { '\\', '/', '?', '*', '[', ']', ':' };
                if (newSheetName.IndexOfAny(invalidChars) >= 0)
                {
                    MessageBox.Show("Tên sheet không được chứa các ký tự: \\ / ? * [ ] :", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Kiểm tra tên sheet đã tồn tại chưa
                foreach (Worksheet ws in activeWorkbook.Worksheets)
                {
                    if (ws.Name.Equals(newSheetName, StringComparison.OrdinalIgnoreCase) && ws != selectedSheet)
                    {
                        MessageBox.Show($"Sheet có tên '{newSheetName}' đã tồn tại. Vui lòng chọn tên khác.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }

                // Đổi tên sheet
                selectedSheet.Name = newSheetName;

                // Cập nhật tất cả hyperlinks trong workbook
                int updatedLinksCount = 0;
                foreach (Worksheet worksheet in activeWorkbook.Worksheets)
                {
                    foreach (Hyperlink hyperlink in worksheet.Hyperlinks)
                    {
                        // Kiểm tra hyperlink có trỏ đến sheet cũ không
                        if (!string.IsNullOrEmpty(hyperlink.Address) &&
                            (hyperlink.Address.StartsWith($"'{oldSheetName}'!") ||
                             hyperlink.Address.StartsWith($"{oldSheetName}!")))
                        {
                            // Cập nhật địa chỉ hyperlink
                            string newAddress = hyperlink.Address.Replace($"'{oldSheetName}'!", $"'{newSheetName}'!")
                                                                  .Replace($"{oldSheetName}!", $"'{newSheetName}'!");
                            hyperlink.Address = newAddress;
                            updatedLinksCount++;
                        }

                        // Cập nhật SubAddress (internal links)
                        if (!string.IsNullOrEmpty(hyperlink.SubAddress) &&
                            (hyperlink.SubAddress.StartsWith($"'{oldSheetName}'!") ||
                             hyperlink.SubAddress.StartsWith($"{oldSheetName}!")))
                        {
                            string newSubAddress = hyperlink.SubAddress.Replace($"'{oldSheetName}'!", $"'{newSheetName}'!")
                                                                       .Replace($"{oldSheetName}!", $"'{newSheetName}'!");
                            hyperlink.SubAddress = newSubAddress;
                            updatedLinksCount++;
                        }
                    }
                }

                // Cập nhật danh sách sheet trong action panel
                if (_actionPanel != null)
                {
                    _actionPanel.listofSheet.DataSource = this.GetListOfSheet();

                    // Tìm và chọn lại sheet với tên mới trong listbox
                    for (int i = 0; i < _actionPanel.listofSheet.Items.Count; i++)
                    {
                        if (_actionPanel.listofSheet.Items[i] is SheetInfo sheetInfo &&
                            sheetInfo.Name == newSheetName)
                        {
                            _actionPanel.listofSheet.SelectedIndex = i;
                            break;
                        }
                    }
                }

                MessageBox.Show($"Đã đổi tên sheet từ '{oldSheetName}' thành '{newSheetName}' thành công.\nĐã cập nhật {updatedLinksCount} hyperlinks.",
                    "Hoàn thành", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi xảy ra khi đổi tên sheet: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// FormatDocument
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormatDocument(object sender, EventArgs e)
        {
            try
            {
                // Lấy Workbook hiện tại
                var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

                if (activeWorkbook != null)
                {
                    // Duyệt qua tất cả các worksheet trong workbook
                    foreach (Worksheet worksheet in activeWorkbook.Worksheets)
                    {
                        // Kích hoạt worksheet
                        worksheet.Activate();

                        // Đặt zoom level về 100%
                        Globals.ThisAddIn.Application.ActiveWindow.Zoom = 100;

                        // Focus vào ô A1
                        worksheet.Range["A1"].Select();

                        // Đảm bảo ô A1 hiển thị ở góc trên bên trái
                        Globals.ThisAddIn.Application.ActiveWindow.ScrollRow = 1;
                        Globals.ThisAddIn.Application.ActiveWindow.ScrollColumn = 1;
                    }

                    // Kích hoạt lại worksheet đầu tiên sau khi format xong
                    if (activeWorkbook.Worksheets.Count > 0)
                    {
                        ((Worksheet)activeWorkbook.Worksheets[1]).Activate();
                        activeWorkbook.Worksheets[1].Range["A1"].Select();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi format document: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// CreateEvidence
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CreateEvidence(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var activeWorkbook = app.ActiveWorkbook;
                var activeSheet = app.ActiveSheet as Worksheet;

                // Kiểm tra workbook và sheet
                if (activeWorkbook == null)
                {
                    MessageBox.Show("Không có workbook nào đang mở. Vui lòng mở một workbook và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                if (activeSheet == null)
                {
                    MessageBox.Show("Không có sheet nào đang được chọn. Vui lòng chọn một sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Kiểm tra cell đang chọn
                Range activeCell = null;
                try { activeCell = app.ActiveCell as Range; } catch { }
                if (activeCell == null)
                {
                    MessageBox.Show("Không có ô nào đang được chọn hoặc lựa chọn không hợp lệ. Vui lòng chọn một ô và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Kiểm tra sheet có bị bảo vệ không
                if (activeSheet.ProtectContents || activeSheet.ProtectDrawingObjects || activeSheet.ProtectScenarios)
                {
                    MessageBox.Show("Sheet đang được bảo vệ. Vui lòng bỏ bảo vệ sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Lấy giá trị của ô hiện tại làm tên sheet mới
                string cellValue = activeCell.Value2 != null ? activeCell.Value2.ToString().Trim() : "";
                if (string.IsNullOrEmpty(cellValue))
                {
                    MessageBox.Show($"Ô hiện tại đang để trống. Vui lòng nhập giá trị vào ô và thử lại.", "Cảnh báo", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Tạo tên sheet mới từ giá trị ô hiện tại
                string newSheetName = cellValue;

                // Kiểm tra sheet đã tồn tại chưa
                Worksheet existingSheet = null;
                foreach (Worksheet ws in activeWorkbook.Worksheets)
                {
                    if (ws.Name == newSheetName)
                    {
                        existingSheet = ws;
                        break;
                    }
                }

                if (existingSheet != null)
                {
                    // Nếu sheet đã tồn tại, chỉ tạo hyperlink đến sheet đó
                    activeSheet.Hyperlinks.Add(activeCell, "", $"'{newSheetName}'!A1", Type.Missing, newSheetName);

                    // Cập nhật selection trong listbox để trỏ đến sheet đã tồn tại
                    if (_actionPanel != null)
                    {
                        for (int i = 0; i < _actionPanel.listofSheet.Items.Count; i++)
                        {
                            if (_actionPanel.listofSheet.Items[i] is SheetInfo sheetInfo &&
                                sheetInfo.Name == newSheetName)
                            {
                                _actionPanel.listofSheet.SelectedIndex = i;
                                break;
                            }
                        }
                    }

                    return;
                }

                // Tạo sheet mới
                Worksheet newWs = activeWorkbook.Worksheets.Add(Type.Missing, activeWorkbook.Worksheets[activeWorkbook.Worksheets.Count]);
                newWs.Name = newSheetName;

                // Đặt độ rộng cột sheet mới
                newWs.Columns.ColumnWidth = 2.17;

                // Tạo hyperlink từ ô hiện tại đến sheet mới
                activeSheet.Hyperlinks.Add(activeCell, "", $"'{newSheetName}'!A1", Type.Missing, newSheetName);

                // Đặt giá trị "Back" vào ô A1 của sheet mới trước khi tạo hyperlink
                newWs.Cells[1, 1].Value2 = "Back";

                // Tạo hyperlink "Back" từ ô A1 của sheet mới về ô gốc
                newWs.Hyperlinks.Add(newWs.Cells[1, 1], "", $"'{activeSheet.Name}'!{activeCell.Address[false, false]}", Type.Missing, "Back");

                // Cập nhật danh sách sheet trong action panel
                if (_actionPanel != null)
                {
                    _actionPanel.listofSheet.DataSource = this.GetListOfSheet();

                    // Tìm và chọn sheet mới tạo trong listbox
                    for (int i = 0; i < _actionPanel.listofSheet.Items.Count; i++)
                    {
                        if (_actionPanel.listofSheet.Items[i] is SheetInfo sheetInfo &&
                            sheetInfo.Name == newSheetName)
                        {
                            _actionPanel.listofSheet.SelectedIndex = i;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi xảy ra khi tạo sheet bằng chứng: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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
            var selectedItem = _actionPanel.listofSheet.SelectedItem as SheetInfo;
            if (selectedItem != null && !string.IsNullOrEmpty(selectedItem.Name))
            {
                // Sử dụng LINQ để tìm worksheet theo tên
                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets
                    .Cast<Worksheet>()
                    .FirstOrDefault(ws => ws.Name == selectedItem.Name);

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
        /// Class để lưu thông tin sheet với màu
        /// </summary>
        public class SheetInfo
        {
            public string Name { get; set; }
            public System.Drawing.Color TabColor { get; set; }
            public bool HasTabColor { get; set; }

            public override string ToString()
            {
                return Name;
            }
        }

        /// <summary>
        /// GetListOfSheet
        /// </summary>
        /// <returns></returns>
        private List<SheetInfo> GetListOfSheet()
        {
            var sheetInfoList = new List<SheetInfo>();

            foreach (Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                var sheetInfo = new SheetInfo
                {
                    Name = sheet.Name,
                    HasTabColor = false,
                    TabColor = System.Drawing.Color.White
                };

                // Kiểm tra xem sheet có màu tab không
                try
                {
                    if (sheet.Tab.Color != null)
                    {
                        // Lấy màu tab của sheet
                        var colorIndex = sheet.Tab.ColorIndex;
                        if (colorIndex != Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone)
                        {
                            sheetInfo.HasTabColor = true;
                            // Chuyển đổi màu Excel sang System.Drawing.Color
                            var excelColor = sheet.Tab.Color;
                            if (excelColor != null)
                            {
                                // Convert Excel color to RGB
                                int rgb = (int)excelColor;
                                sheetInfo.TabColor = System.Drawing.Color.FromArgb(
                                    rgb & 0xFF,           // Red
                                    (rgb >> 8) & 0xFF,    // Green
                                    (rgb >> 16) & 0xFF    // Blue
                                );
                            }
                        }
                    }
                }
                catch
                {
                    // Nếu có lỗi khi lấy màu, sử dụng màu mặc định
                    sheetInfo.HasTabColor = false;
                    sheetInfo.TabColor = System.Drawing.Color.White;
                }

                sheetInfoList.Add(sheetInfo);
            }

            return sheetInfoList;
        }

        /// <summary>
        /// FindIndexOfSelectedSheet
        /// </summary>
        /// <returns></returns>
        private int FindIndexOfSelectedSheet()
        {
            if (_actionPanel?.listofSheet?.Items == null) return -1;

            var currentSheetName = Globals.ThisAddIn.Application.ActiveWorkbook?.ActiveSheet?.Name;
            if (string.IsNullOrEmpty(currentSheetName)) return -1;

            for (int i = 0; i < _actionPanel.listofSheet.Items.Count; i++)
            {
                if (_actionPanel.listofSheet.Items[i] is SheetInfo sheetInfo &&
                    sheetInfo.Name == currentSheetName)
                {
                    return i;
                }
            }
            return -1;
        }
        #endregion
    }
}