using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Windows;

namespace ExcelCustomAddin
{
    /// <summary>
    /// SheetManagementService - Quản lý các chức năng liên quan đến sheet
    ///
    /// Chức năng chính:
    /// - Đổi tên sheet với validation và hyperlink updates
    /// - Quản lý danh sách sheet với pin/unpin functionality
    /// - Activate sheet từ ActionPanel selection
    /// - Format document (zoom và scroll settings)
    /// - Debounced selection changes để tránh recursive updates
    ///
    /// Tính năng đặc biệt:
    /// - Pin sheets: Giữ sheets quan trọng ở đầu danh sách
    /// - Hyperlink updates: Tự động cập nhật khi đổi tên sheet
    /// - Debounce: Tránh recursive activation khi selection changes
    /// - Sheet validation: Kiểm tra tên hợp lệ và không trùng lặp
    ///
    /// Tác giả: lam.pt
    /// Ngày tạo: 2025
    /// </summary>
    public class SheetManagementService
    {
        #region Fields

        private readonly ThisAddIn _addIn;

        // Debounce cho selection changes để tránh recursive updates
        private DateTime lastSelectionChangeTime = DateTime.MinValue;

        #endregion

        #region Constructor

        /// <summary>
        /// Khởi tạo SheetManagementService
        ///
        /// </summary>
        /// <param name="addIn">Instance của ThisAddIn chính</param>
        public SheetManagementService(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        #endregion

        #region Public Interface

        /// <summary>
        /// ChangeSheetName - Đổi tên sheet được chọn từ ActionPanel
        ///
        /// Quy trình:
        /// 1. Validate workbook và selection từ ActionPanel
        /// 2. Hiển thị input dialog để nhập tên mới
        /// 3. Validate tên mới (độ dài, ký tự hợp lệ, không trùng)
        /// 4. Đổi tên sheet và cập nhật tất cả hyperlinks
        /// 5. Refresh ActionPanel UI
        ///
        /// </summary>
        public void ChangeSheetName()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var activeWorkbook = app.ActiveWorkbook;

                // Kiểm tra workbook
                if (activeWorkbook == null)
                {
                    Logger.Error("Không có workbook nào đang mở trong RenameWorksheet");
                    MessageBox.Show("Không có workbook nào đang mở. Vui lòng mở một workbook và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Lấy tên sheet từ ListView thay vì active sheet
                if (_addIn._actionPanel.listofSheet.SelectedItems.Count == 0)
                {
                    Logger.Error("Không có sheet nào được chọn từ danh sách để đổi tên");
                    MessageBox.Show("Vui lòng chọn một sheet từ danh sách để đổi tên.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var selectedItem = _addIn._actionPanel.listofSheet.SelectedItems[0].Tag as SheetInfo;
                if (selectedItem == null || string.IsNullOrEmpty(selectedItem.Name))
                {
                    Logger.Error("Không thể lấy thông tin sheet được chọn");
                    MessageBox.Show("Không thể lấy thông tin sheet được chọn.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    Logger.Error($"Không tìm thấy sheet có tên '{selectedSheetName}'");
                    MessageBox.Show($"Không tìm thấy sheet có tên '{selectedSheetName}'.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Kiểm tra sheet có bị bảo vệ không
                if (selectedSheet.ProtectContents || selectedSheet.ProtectDrawingObjects || selectedSheet.ProtectScenarios)
                {
                    Logger.Error($"Sheet '{selectedSheetName}' đang được bảo vệ");
                    MessageBox.Show($"Sheet '{selectedSheetName}' đang được bảo vệ. Vui lòng bỏ bảo vệ sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string oldSheetName = selectedSheet.Name;

                // Hiển thị dialog để nhập tên mới
                string newSheetName = "";

                // Sử dụng Application.InputBox của Excel thay vì Microsoft.VisualBasic
                object result = app.InputBox(
                    $"Nhập tên mới cho sheet '{oldSheetName}':",
                    "Đổi tên Sheet",
                    oldSheetName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2 // Type 2 = text
                );

                // Kiểm tra user có hủy không
                // Khi user nhấn Cancel, Excel InputBox trả về false (boolean)
                // Khi user để trống và OK, trả về empty string
                // Khi user nhập dữ liệu hợp lệ, trả về string
                if (result == null || result is bool)
                {
                    return; // User canceled
                }

                string resultString = result.ToString().Trim();
                if (string.IsNullOrEmpty(resultString))
                {
                    return; // User provided empty input
                }

                newSheetName = resultString;

                // Kiểm tra user có nhập tên mới không
                if (string.IsNullOrWhiteSpace(newSheetName) || newSheetName == oldSheetName)
                {
                    return; // User hủy hoặc không thay đổi
                }

                // Kiểm tra tên sheet mới có hợp lệ không
                if (newSheetName.Length > 31)
                {
                    Logger.Error($"Tên sheet '{newSheetName}' vượt quá 31 ký tự");
                    MessageBox.Show("Tên sheet không được vượt quá 31 ký tự.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Kiểm tra ký tự không hợp lệ
                char[] invalidChars = { '\\', '/', '?', '*', '[', ']', ':' };
                if (newSheetName.IndexOfAny(invalidChars) >= 0)
                {
                    Logger.Error($"Tên sheet '{newSheetName}' chứa ký tự không hợp lệ");
                    MessageBox.Show("Tên sheet không được chứa các ký tự: \\ / ? * [ ] :", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Kiểm tra tên sheet đã tồn tại chưa
                foreach (Worksheet ws in activeWorkbook.Worksheets)
                {
                    if (ws.Name.Equals(newSheetName, StringComparison.OrdinalIgnoreCase) && ws != selectedSheet)
                    {
                        Logger.Error($"Sheet có tên '{newSheetName}' đã tồn tại");
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
                if (_addIn._actionPanel != null)
                {
                    _addIn._actionPanel.BindSheetList(this.GetListOfSheet(), newSheetName);
                }

                Logger.Info($"Successfully renamed sheet from '{oldSheetName}' to '{newSheetName}', updated {updatedLinksCount} hyperlinks");
            }
            catch (Exception ex)
            {
                Logger.Error($"Có lỗi xảy ra khi đổi tên sheet: {ex.Message}", ex);
                MessageBox.Show($"Có lỗi xảy ra khi đổi tên sheet: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// FormatDocument - Định dạng toàn bộ document với zoom và scroll settings
        ///
        /// Áp dụng cho tất cả worksheets:
        /// - Set zoom level từ config
        /// - Scroll to top-left (A1)
        /// - Activate từng sheet để apply settings
        ///
        /// </summary>
        public void FormatDocument()
        {
            try
            {
                var config = SheetConfigManager.GetGeneralConfig();

                // Lấy Workbook hiện tại
                var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

                if (activeWorkbook != null)
                {
                    // Duyệt qua tất cả các worksheet trong workbook
                    foreach (Worksheet worksheet in activeWorkbook.Worksheets)
                    {
                        // Kích hoạt worksheet
                        worksheet.Activate();

                        // Đặt zoom level từ config
                        Globals.ThisAddIn.Application.ActiveWindow.Zoom = config.WindowZoom;

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
                Logger.Error($"Lỗi khi format document: {ex.Message}", ex);
                MessageBox.Show($"Lỗi khi format document: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// PinSheet - Toggle pin status của sheet
        ///
        /// </summary>
        /// <param name="e">Event args chứa tên sheet cần pin/unpin</param>
        public void PinSheet(ActionPanelControl.PinSheetEventArgs e)
        {
            try
            {
                var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (activeWorkbook == null)
                {
                    Logger.Error("Không có workbook nào đang mở trong PinSheet");
                    MessageBox.Show("Không có workbook nào đang mở.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string workbookName = activeWorkbook.Name;
                TogglePinSheet(workbookName, e.SheetName);
            }
            catch (Exception ex)
            {
                Logger.Error($"Có lỗi xảy ra khi ghim/bỏ ghim sheet: {ex.Message}", ex);
                MessageBox.Show($"Có lỗi xảy ra khi ghim/bỏ ghim sheet: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Sheet List Management

        /// <summary>
        /// Build and return the list of sheet info for the active workbook
        /// Bao gồm thông tin pin status và tab color
        ///
        /// </summary>
        /// <returns>Danh sách SheetInfo được sort với pinned sheets ở đầu</returns>
        public System.Collections.Generic.List<SheetInfo> GetListOfSheet()
        {
            var sheetInfoList = new System.Collections.Generic.List<SheetInfo>();
            var workbookName = Globals.ThisAddIn.Application.ActiveWorkbook?.Name;

            foreach (Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Sheets)
            {
                var sheetInfo = new SheetInfo
                {
                    Name = sheet.Name,
                    HasTabColor = false,
                    TabColor = System.Drawing.Color.White,
                    IsPinned = workbookName != null && IsSheetPinned(workbookName, sheet.Name)
                };

                try
                {
                    if (sheet.Tab.Color != null)
                    {
                        var colorIndex = sheet.Tab.ColorIndex;
                        if (colorIndex != Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone)
                        {
                            sheetInfo.HasTabColor = true;
                            var excelColor = sheet.Tab.Color;
                            if (excelColor != null)
                            {
                                int rgb = (int)excelColor;
                                sheetInfo.TabColor = System.Drawing.Color.FromArgb(
                                  rgb & 0xFF,
                                  (rgb >> 8) & 0xFF,
                                  (rgb >> 16) & 0xFF);
                            }
                        }
                    }
                }
                catch
                {
                    sheetInfo.HasTabColor = false;
                    sheetInfo.TabColor = System.Drawing.Color.White;
                }

                sheetInfoList.Add(sheetInfo);
            }

            return sheetInfoList.OrderByDescending(s => s.IsPinned).ToList();
        }

        /// <summary>
        /// Find index in the ActionPanel list for the currently active sheet
        ///
        /// </summary>
        /// <returns>Index của active sheet trong list, -1 nếu không tìm thấy</returns>
        public int FindIndexOfSelectedSheet()
        {
            if (_addIn._actionPanel?.listofSheet?.Items == null) return -1;

            var currentSheetName = Globals.ThisAddIn.Application.ActiveWorkbook?.ActiveSheet?.Name;
            if (string.IsNullOrEmpty(currentSheetName)) return -1;

            for (int i = 0; i < _addIn._actionPanel.listofSheet.Items.Count; i++)
            {
                var sheetInfo = _addIn._actionPanel.listofSheet.Items[i].Tag as SheetInfo;
                if (sheetInfo != null && sheetInfo.Name == currentSheetName)
                {
                    return i;
                }
            }
            return -1;
        }

        #endregion

        #region Sheet Activation

        /// <summary>
        /// Activate the sheet selected from ActionPanel list
        /// Tìm và activate sheet theo tên từ selection
        ///
        /// </summary>
        public void SetActiveSheet()
        {
            try
            {
                if (_addIn._actionPanel?.listofSheet?.SelectedItems != null &&
                    _addIn._actionPanel.listofSheet.SelectedItems.Count > 0)
                {
                    var selectedItem = _addIn._actionPanel.listofSheet.SelectedItems[0].Tag as SheetInfo;
                    if (selectedItem != null && !string.IsNullOrEmpty(selectedItem.Name))
                    {
                        var activeWorkbook = Globals.ThisAddIn.Application?.ActiveWorkbook;
                        if (activeWorkbook != null)
                        {
                            Worksheet sheet = activeWorkbook.Sheets
                                .Cast<Worksheet>()
                                .FirstOrDefault(ws => ws.Name == selectedItem.Name);

                            if (sheet != null)
                            {
                                sheet.Activate();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error in SetActiveSheet: {ex.Message}");
            }
        }

        /// <summary>
        /// Selection changed handler for the ActionPanel list - includes debounce
        /// Sử dụng debounce để tránh recursive activation loops
        ///
        /// </summary>
        /// <param name="sender">Event sender</param>
        /// <param name="e">Event args</param>
        public void ListOfSheet_SelectionChanged(object sender, EventArgs e)
        {
            if (_addIn.IsSheetActivating)
            {
                return;
            }

            lastSelectionChangeTime = DateTime.Now;
            var currentTime = lastSelectionChangeTime;
            System.Threading.Tasks.Task.Delay(300).ContinueWith(t =>
            {
                if (currentTime == lastSelectionChangeTime)
                {
                    try
                    {
                        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
                        SetActiveSheet();
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"Exception in delayed SetActiveSheet: {ex.Message}");
                    }
                }
            });
        }

        #endregion

        #region Pin Management

        /// <summary>
        /// Toggle pin status của sheet
        /// Thread-safe với lock để tránh race conditions
        ///
        /// </summary>
        /// <param name="workbookName">Tên workbook</param>
        /// <param name="sheetName">Tên sheet cần toggle</param>
        public void TogglePinSheet(string workbookName, string sheetName)
        {
            TogglePinSheetPrivate(workbookName, sheetName);
        }

        /// <summary>
        /// Private implementation của TogglePinSheet với thread safety
        ///
        /// </summary>
        /// <param name="workbookName">Tên workbook</param>
        /// <param name="sheetName">Tên sheet cần toggle</param>
        private void TogglePinSheetPrivate(string workbookName, string sheetName)
        {
            lock (ThisAddIn._lockObject)
            {
                if (!ThisAddIn.PinnedSheets.ContainsKey(workbookName))
                {
                    ThisAddIn.PinnedSheets[workbookName] = new System.Collections.Generic.HashSet<string>();
                }

                var pinnedSheets = ThisAddIn.PinnedSheets[workbookName];

                if (pinnedSheets.Contains(sheetName))
                {
                    // Bỏ ghim
                    pinnedSheets.Remove(sheetName);
                    Logger.Info($"Unpinned sheet '{sheetName}' in workbook '{workbookName}'");
                }
                else
                {
                    // Ghim
                    pinnedSheets.Add(sheetName);
                    Logger.Info($"Pinned sheet '{sheetName}' in workbook '{workbookName}'");
                }

                // Cập nhật UI
                if (_addIn._actionPanel != null)
                {
                    _addIn._actionPanel.BindSheetList(this.GetListOfSheet());
                }
            }
        }

        /// <summary>
        /// Kiểm tra sheet có được pin không
        ///
        /// </summary>
        /// <param name="workbookName">Tên workbook</param>
        /// <param name="sheetName">Tên sheet cần kiểm tra</param>
        /// <returns>true nếu sheet được pin, false nếu không</returns>
        public bool IsSheetPinned(string workbookName, string sheetName)
        {
            return ThisAddIn.PinnedSheets.ContainsKey(workbookName) && ThisAddIn.PinnedSheets[workbookName].Contains(sheetName);
        }

        #endregion
    }
}