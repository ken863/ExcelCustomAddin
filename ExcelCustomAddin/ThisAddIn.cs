namespace ExcelCustomAddin
{
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
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
        /// Dispatcher
        /// </summary>
        private Dispatcher _dispatcher;

        /// <summary>
        /// Kiểm tra sheet đang trong quá trình kích hoạt 
        /// để ngăn chặn sự kiện update danh sách sheet khi SelectedIndexChanged và SheetActivate
        /// </summary>
        private bool IsSheetActivating { get; set; } = false;

        /// <summary>
        /// Lưu trữ danh sách các sheet được pin theo workbook
        /// </summary>
        private static Dictionary<string, HashSet<string>> PinnedSheets
            = new Dictionary<string, HashSet<string>>();

        /// <summary>
        /// Lưu trữ workbook đã được tạo action panel để tránh tạo trùng lặp
        /// </summary>
        private static HashSet<string> CreatedActionPanes
            = new HashSet<string>();

        /// <summary>
        /// Lock object để đảm bảo thread safety
        /// </summary>
        private static readonly object _lockObject = new object();        /// <summary>
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
                    Globals.ThisAddIn.Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
                    Globals.ThisAddIn.Application.WorkbookAfterSave -= Application_WorkbookAfterSave;
                    Globals.ThisAddIn.Application.SheetActivate -= Application_SheetActivate;
                }

                // Hủy đăng ký action panel events
                if (_actionPanel != null)
                {
                    _actionPanel.CreateEvidenceEvent -= this.CreateEvidence;
                    _actionPanel.FormatDocumentEvent -= this.FormatDocument;
                    _actionPanel.ChangeSheetNameEvent -= this.ChangeSheetName;
                    _actionPanel.PinSheetEvent -= this.PinSheet;
                    _actionPanel.InsertMultipleImagesEvent -= this.InsertMultipleImages;
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
        /// Xử lý sự kiện chèn nhiều hình ảnh
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InsertMultipleImages(object sender, EventArgs e)
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

                // Đường dẫn thư mục chứa hình ảnh
                string folderPath = _actionPanel.txtImagePath.Text.Trim();

                // Tạo thư mục nếu chưa tồn tại
                if (!System.IO.Directory.Exists(folderPath))
                {
                    System.IO.Directory.CreateDirectory(folderPath);
                    MessageBox.Show($"Đã tạo thư mục '{folderPath}'. Vui lòng thêm hình ảnh vào thư mục này và thử lại.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Lấy danh sách file hình ảnh
                string[] imageExtensions = { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.gif", "*.tiff" };
                var imageFiles = new List<string>();

                foreach (string extension in imageExtensions)
                {
                    var files = System.IO.Directory.GetFiles(folderPath, extension, System.IO.SearchOption.TopDirectoryOnly);
                    imageFiles.AddRange(files);
                }

                if (imageFiles.Count == 0)
                {
                    MessageBox.Show($"Không tìm thấy file hình ảnh nào trong thư mục '{folderPath}'.\nCác định dạng được hỗ trợ: JPG, JPEG, PNG, BMP, GIF, TIFF", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Lấy vị trí bắt đầu từ cell hiện tại
                double topLocation = (double)activeCell.Top;
                double leftLocation = (double)activeCell.Left;
                double resizeRate = (double)(_actionPanel.numScalePercent.Value / 100); // Tỷ lệ thu nhỏ hình ảnh
                int insertedCount = 0;
                int errorCount = 0;

                // Chèn từng hình ảnh
                foreach (string imagePath in imageFiles)
                {
                    try
                    {
                        // Chèn hình ảnh vào sheet
                        var shape = activeSheet.Shapes.AddPicture(
                            imagePath,
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue,
                            (float)leftLocation,
                            (float)topLocation,
                            -1, // Width - tự động
                            -1  // Height - tự động
                        );

                        // Điều chỉnh kích thước hình ảnh
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                        shape.Height = (float)(shape.Height * resizeRate);

                        // Cập nhật vị trí cho hình ảnh tiếp theo
                        topLocation += shape.Height + activeCell.Height;

                        insertedCount++;

                        // Xóa file sau khi chèn thành công
                        try
                        {
                            // Xóa thuộc tính readonly nếu có
                            System.IO.File.SetAttributes(imagePath, System.IO.FileAttributes.Normal);
                            System.IO.File.Delete(imagePath);
                        }
                        catch (Exception deleteEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"Không thể xóa file {imagePath}: {deleteEx.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        System.Diagnostics.Debug.WriteLine($"Lỗi khi chèn hình ảnh {imagePath}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi xảy ra khi chèn hình ảnh: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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
            Globals.ThisAddIn.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
            Globals.ThisAddIn.Application.WorkbookAfterSave += Application_WorkbookAfterSave;
            Globals.ThisAddIn.Application.SheetActivate += Application_SheetActivate;

            // Tạo ActionPane cho workbook hiện tại (nếu có) với delay để tránh trùng lặp
            if (this.Application.ActiveWorkbook != null)
            {
                // Sử dụng timer để đảm bảo chỉ tạo 1 lần sau khi startup xong
                var timer = new System.Windows.Forms.Timer();
                timer.Interval = 500; // 500ms delay
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    this.CreateActionsPane(this.Application.ActiveWorkbook);
                };
                timer.Start();
            }
        }

        #region "Quản lý ActionPane"
        /// <summary>
        /// Application_NewWorkbook
        /// </summary>
        /// <param name="Wb"></param>
        void Application_NewWorkbook(Workbook Wb)
        {
            System.Diagnostics.Debug.WriteLine($"Application_NewWorkbook called for: {Wb?.Name}");
            this.CreateActionsPane(Wb);
        }

        /// <summary>
        /// Application_WorkbookActivate
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookOpen(Workbook Wb)
        {
            System.Diagnostics.Debug.WriteLine($"Application_WorkbookOpen called for: {Wb?.Name}");
            this.CreateActionsPane(Wb);
        }

        /// <summary>
        /// Application_WorkbookActivate
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookActive(Workbook Wb)
        {
            // Khi activate workbook, chỉ cập nhật action panel nếu đã tồn tại
            // Không tạo mới để tránh trùng lặp với Open/New events
            if (Wb != null && CreatedActionPanes.Contains(Wb.Name))
            {
                // Chỉ cập nhật nếu đã có action panel
                if (_actionPanel != null)
                {
                    var currentSheetName = Wb.ActiveSheet?.Name;
                    _actionPanel.BindSheetList(this.GetListOfSheet(), currentSheetName);
                }
            }
        }

        /// <summary>
        /// Application_WorkbookBeforeClose
        /// </summary>
        /// <param name="Wb"></param>
        /// <param name="Cancel"></param>
        private void Application_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            if (Wb != null)
            {
                string workbookKey = Wb.Name;

                // Xóa workbook khỏi danh sách đã tạo action panel
                if (CreatedActionPanes.Contains(workbookKey))
                {
                    CreatedActionPanes.Remove(workbookKey);
                }

                // Xóa pinned sheets của workbook này
                if (PinnedSheets.ContainsKey(workbookKey))
                {
                    PinnedSheets.Remove(workbookKey);
                }
            }
        }

        /// <summary>
        /// Application_WorkbookAfterSave - Cập nhật file path sau khi lưu
        /// </summary>
        /// <param name="Wb"></param>
        /// <param name="Success"></param>
        private void Application_WorkbookAfterSave(Workbook Wb, bool Success)
        {
            if (Wb != null && Success && _actionPanel != null)
            {
                // Cập nhật file path display sau khi workbook được lưu thành công
                _actionPanel.RefreshFilePathDisplay();
                System.Diagnostics.Debug.WriteLine($"File path refreshed after save for: {Wb.Name}");
            }
        }

        private void CreateActionsPane(Workbook Wb)
        {
            if (Wb != null)
            {
                string workbookKey = Wb.Name;

                lock (_lockObject)
                {
                    // Debug logging
                    System.Diagnostics.Debug.WriteLine($"CreateActionsPane called for: {workbookKey}");

                    // Kiểm tra xem action panel đã được tạo cho workbook này chưa
                    if (CreatedActionPanes.Contains(workbookKey))
                    {
                        System.Diagnostics.Debug.WriteLine($"Action panel already exists for: {workbookKey}, updating only");
                        // Nếu đã tạo rồi, chỉ cần cập nhật danh sách sheet
                        if (_actionPanel != null && myCustomTaskPane != null)
                        {
                            // Đảm bảo task pane đang active cho workbook này
                            var currentTaskPane = TaskPaneManager.GetTaskPane(workbookKey, "WORKSHEET TOOLS", null);
                            if (currentTaskPane != null)
                            {
                                myCustomTaskPane = currentTaskPane;
                                _actionPanel = (ActionPanelControl)myCustomTaskPane.Control;

                                var currentSheetName = Wb.ActiveSheet?.Name;
                                _actionPanel.BindSheetList(this.GetListOfSheet(), currentSheetName);
                            }
                        }
                        return;
                    }

                    System.Diagnostics.Debug.WriteLine($"Creating new action panel for: {workbookKey}");

                    // Get Active ActionsPanel
                    myCustomTaskPane = TaskPaneManager.GetTaskPane(Wb.Name, "WORKSHEET TOOLS", () => new ActionPanelControl());
                    _actionPanel = (ActionPanelControl)myCustomTaskPane?.Control;

                    if (_actionPanel != null)
                    {
                        // Hủy đăng ký các event cũ trước khi đăng ký mới để tránh đăng ký trùng lặp
                        _actionPanel.CreateEvidenceEvent -= this.CreateEvidence;
                        _actionPanel.FormatDocumentEvent -= this.FormatDocument;
                        _actionPanel.ChangeSheetNameEvent -= this.ChangeSheetName;
                        _actionPanel.InsertMultipleImagesEvent -= this.InsertMultipleImages;
                        _actionPanel.PinSheetEvent -= this.PinSheet;
                        _actionPanel.listofSheet.SelectedIndexChanged -= this.ListOfSheet_SelectionChanged;

                        // Đăng ký các event mới
                        _actionPanel.CreateEvidenceEvent += this.CreateEvidence;
                        _actionPanel.FormatDocumentEvent += this.FormatDocument;
                        _actionPanel.ChangeSheetNameEvent += this.ChangeSheetName;
                        _actionPanel.InsertMultipleImagesEvent += this.InsertMultipleImages;
                        _actionPanel.PinSheetEvent += this.PinSheet;
                        _actionPanel.listofSheet.SelectedIndexChanged += this.ListOfSheet_SelectionChanged;

                        // Cập nhật danh sách sheet và chọn sheet hiện tại khi tạo ActionPane
                        var currentSheetName = Wb.ActiveSheet?.Name;
                        _actionPanel.BindSheetList(this.GetListOfSheet(), currentSheetName);

                        // *** THÊM DÒNG NÀY: Tự động hiển thị Action Panel khi workbook được mở ***
                        myCustomTaskPane.Visible = true;

                        // Tùy chọn: Đặt độ rộng mặc định cho task pane (tuỳ chỉnh theo nhu cầu)
                        myCustomTaskPane.Width = 300;

                        // Đánh dấu workbook này đã được tạo action panel
                        CreatedActionPanes.Add(workbookKey);
                        System.Diagnostics.Debug.WriteLine($"Action panel created and marked for: {workbookKey}");
                    }
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

                // Lấy tên sheet từ ListView thay vì active sheet
                if (_actionPanel.listofSheet.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Vui lòng chọn một sheet từ danh sách để đổi tên.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var selectedItem = _actionPanel.listofSheet.SelectedItems[0].Tag as SheetInfo;
                if (selectedItem == null || string.IsNullOrEmpty(selectedItem.Name))
                {
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
                    _actionPanel.BindSheetList(this.GetListOfSheet(), newSheetName);
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
        /// PinSheet - Toggle pin status của sheet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void PinSheet(object sender, ActionPanelControl.PinSheetEventArgs e)
        {
            try
            {
                var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (activeWorkbook == null)
                {
                    MessageBox.Show("Không có workbook nào đang mở.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string workbookName = activeWorkbook.Name;
                TogglePinSheet(workbookName, e.SheetName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi xảy ra khi ghim/bỏ ghim sheet: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
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

                    // Cập nhật selection trong ListView để trỏ đến sheet đã tồn tại
                    if (_actionPanel != null)
                    {
                        _actionPanel.BindSheetList(this.GetListOfSheet(), newSheetName);
                    }

                    return;
                }

                // Tạo sheet mới
                Worksheet newWs = activeWorkbook.Worksheets.Add(Type.Missing, activeWorkbook.Worksheets[activeWorkbook.Worksheets.Count]);
                newWs.Name = newSheetName;

                // Đặt độ rộng cột sheet mới
                newWs.Columns.ColumnWidth = 2.17;
                newWs.Rows.RowHeight = 12.75; // Giảm chiều cao dòng để fit 48 dòng trên Windows

                // Thiết lập font chữ cho toàn bộ sheet
                newWs.Cells.Font.Name = "MS PGothic";
                newWs.Cells.Font.Size = 9; // Giảm font size để fit 48 dòng trên Windows

                // Thiết lập used range tới cột BC (cột 55)
                // Đặt giá trị vào ô BC1 để mở rộng used range
                newWs.Cells[1, 55].Value2 = " ";

                // Thiết lập trang in với định hướng ngang và lề trang tới cột BC
                newWs.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                newWs.PageSetup.PaperSize = XlPaperSize.xlPaperA4; // Thiết lập kích cỡ giấy A4
                newWs.PageSetup.PrintArea = "$A$1:$BC$48"; // Thiết lập vùng in từ A1 đến BC48
                newWs.PageSetup.FitToPagesWide = 1; // Fit tất cả cột vào 1 trang theo chiều rộng
                newWs.PageSetup.FitToPagesTall = 1; // Fit 48 dòng vào 1 trang theo chiều cao
                newWs.PageSetup.Zoom = false; // Tắt zoom để sử dụng FitToPages

                // Thiết lập lề trang tối ưu cho Windows (đơn vị: inches)
                newWs.PageSetup.LeftMargin = newWs.Application.InchesToPoints(0.2);   // Lề trái nhỏ hơn
                newWs.PageSetup.RightMargin = newWs.Application.InchesToPoints(0.2);  // Lề phải nhỏ hơn
                newWs.PageSetup.TopMargin = newWs.Application.InchesToPoints(0.2);    // Lề trên nhỏ hơn
                newWs.PageSetup.BottomMargin = newWs.Application.InchesToPoints(0.2); // Lề dưới nhỏ hơn
                newWs.PageSetup.HeaderMargin = newWs.Application.InchesToPoints(0.05); // Lề header nhỏ hơn
                newWs.PageSetup.FooterMargin = newWs.Application.InchesToPoints(0.05); // Lề footer nhỏ hơn

                // Thiết lập view mode thành Page Break Preview
                try
                {
                    var window = newWs.Application.ActiveWindow;
                    if (window != null)
                    {
                        window.View = XlWindowView.xlPageBreakPreview;
                        // Thiết lập zoom về 100%
                        window.Zoom = 100;
                    }
                }
                catch (Exception viewEx)
                {
                    // Log error nhưng không làm gián đoạn quá trình tạo sheet
                    System.Diagnostics.Debug.WriteLine($"Error setting page break preview or zoom: {viewEx.Message}");
                }

                // Tạo hyperlink từ ô hiện tại đến sheet mới
                activeSheet.Hyperlinks.Add(activeCell, "", $"'{newSheetName}'!A1", Type.Missing, newSheetName);

                // Đặt giá trị "Back" vào ô A1 của sheet mới trước khi tạo hyperlink
                newWs.Cells[1, 1].Value2 = "Back";

                // Tạo hyperlink "Back" từ ô A1 của sheet mới về ô gốc
                newWs.Hyperlinks.Add(newWs.Cells[1, 1], "", $"'{activeSheet.Name}'!{activeCell.Address[false, false]}", Type.Missing, "戻る");

                // Cập nhật danh sách sheet trong action panel
                if (_actionPanel != null)
                {
                    _actionPanel.BindSheetList(this.GetListOfSheet(), newSheetName);
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
        /// Last selection change time để debounce
        /// </summary>
        private DateTime lastSelectionChangeTime = DateTime.MinValue;

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

            // Debounce để tránh gọi quá nhiều lần
            lastSelectionChangeTime = DateTime.Now;

            // Sử dụng Task.Delay để debounce
            var currentTime = lastSelectionChangeTime;
            System.Threading.Tasks.Task.Delay(300).ContinueWith(t =>
            {
                if (currentTime == lastSelectionChangeTime)
                {
                    // Chỉ thực hiện nếu không có selection change mới
                    try
                    {
                        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
                        this.SetActiveSheet();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Exception in delayed SetActiveSheet: {ex.Message}");
                    }
                }
            });
        }

        /// <summary>
        /// SetActiveSheet
        /// </summary>
        private void SetActiveSheet()
        {
            try
            {
                if (_actionPanel?.listofSheet?.SelectedItems != null &&
                    _actionPanel.listofSheet.SelectedItems.Count > 0)
                {
                    var selectedItem = _actionPanel.listofSheet.SelectedItems[0].Tag as SheetInfo;
                    if (selectedItem != null && !string.IsNullOrEmpty(selectedItem.Name))
                    {
                        var activeWorkbook = Globals.ThisAddIn.Application?.ActiveWorkbook;
                        if (activeWorkbook != null)
                        {
                            // Sử dụng LINQ để tìm worksheet theo tên
                            Worksheet sheet = activeWorkbook.Sheets
                                .Cast<Worksheet>()
                                .FirstOrDefault(ws => ws.Name == selectedItem.Name);

                            if (sheet != null)
                            {
                                // Đặt worksheet này là active sheet
                                sheet.Activate();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in SetActiveSheet: {ex.Message}");
            }
        }

        /// <summary>
        /// Application_SheetActivate
        /// </summary>
        /// <param name="Sh"></param>
        private void Application_SheetActivate(object Sh)
        {
            try
            {
                this.IsSheetActivating = true;
                var activeSheetName = Globals.ThisAddIn.Application?.ActiveWorkbook?.ActiveSheet?.Name;
                if (_actionPanel != null)
                {
                    _actionPanel.BindSheetList(this.GetListOfSheet(), activeSheetName);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Application_SheetActivate: {ex.Message}");
            }
            finally
            {
                this.IsSheetActivating = false;
            }
        }

        /// <summary>
        /// Class để lưu thông tin sheet với màu và trạng thái pin
        /// </summary>
        public class SheetInfo
        {
            public string Name { get; set; }
            public System.Drawing.Color TabColor { get; set; }
            public bool HasTabColor { get; set; }
            public bool IsPinned { get; set; } = false;

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

            // Sắp xếp: sheet được pin lên đầu, sau đó theo thứ tự bình thường
            return sheetInfoList.OrderByDescending(s => s.IsPinned).ToList();
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
                var sheetInfo = _actionPanel.listofSheet.Items[i].Tag as SheetInfo;
                if (sheetInfo != null && sheetInfo.Name == currentSheetName)
                {
                    return i;
                }
            }
            return -1;
        }
        #endregion

        #region "Pin Sheet Functionality"
        /// <summary>
        /// Toggle pin status của sheet
        /// </summary>
        /// <param name="workbookName"></param>
        /// <param name="sheetName"></param>
        public void TogglePinSheet(string workbookName, string sheetName)
        {
            if (!PinnedSheets.ContainsKey(workbookName))
            {
                PinnedSheets[workbookName] = new HashSet<string>();
            }

            if (PinnedSheets[workbookName].Contains(sheetName))
            {
                PinnedSheets[workbookName].Remove(sheetName);
            }
            else
            {
                PinnedSheets[workbookName].Add(sheetName);
            }

            // Cập nhật lại danh sách sheet
            if (_actionPanel != null)
            {
                var currentSheetName = Globals.ThisAddIn.Application.ActiveWorkbook?.ActiveSheet?.Name;
                _actionPanel.BindSheetList(this.GetListOfSheet(), currentSheetName);
            }
        }

        /// <summary>
        /// Kiểm tra xem sheet có được pin không
        /// </summary>
        /// <param name="workbookName"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public bool IsSheetPinned(string workbookName, string sheetName)
        {
            return PinnedSheets.ContainsKey(workbookName) && PinnedSheets[workbookName].Contains(sheetName);
        }
        #endregion
    }
}