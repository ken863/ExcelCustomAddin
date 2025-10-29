namespace ExcelCustomAddin
{
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
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
        private static readonly object _lockObject = new object();

        // Column used as the page-break / right-most printed column for Evidence sheets
        private static string PAGE_BREAK_COLUMN_NAME => SheetConfigManager.GetGeneralConfig().PageBreakColumnName;
        // Default font used for Evidence sheets and hyperlink cells
        private static string EVIDENCE_FONT_NAME => SheetConfigManager.GetGeneralConfig().EvidenceFontName;
        // Last row index for print area in Evidence sheets
        private static int PRINT_AREA_LAST_ROW_IDX => SheetConfigManager.GetGeneralConfig().PrintAreaLastRowIdx;

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
                    _actionPanel.FormatImagesEvent -= this.FormatImages;
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
                Logger.Error($"Error during shutdown: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Apply picture style to shape: no fill, no line, shadow, no reflection, no glow
        /// </summary>
        /// <param name="shape">The shape to apply style to</param>
        private void ApplyPictureStyleToShape(Microsoft.Office.Interop.Excel.Shape shape)
        {
            shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            shape.Shadow.Style = Microsoft.Office.Core.MsoShadowStyle.msoShadowStyleOuterShadow;
            shape.Shadow.Type = Microsoft.Office.Core.MsoShadowType.msoShadow21;
            shape.Shadow.ForeColor.RGB = 0; // Black
            shape.Shadow.Transparency = 0.3f;
            shape.Shadow.Size = 100;
            shape.Shadow.Blur = 15;
            shape.Shadow.OffsetX = 0;
            shape.Shadow.OffsetY = 0;
            shape.Reflection.Type = Microsoft.Office.Core.MsoReflectionType.msoReflectionTypeNone;
            shape.Glow.Radius = 0;
            shape.SoftEdge.Radius = 0;
        }

        /// <summary>
        /// FormatImages
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormatImages(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;

                // Lấy các shape đã chọn
                ShapeRange selectedShapes = null;
                try
                {
                    selectedShapes = app.Selection.ShapeRange;
                }
                catch
                {
                    Logger.Error("Không có hình ảnh nào được chọn. Vui lòng chọn các hình ảnh và thử lại.");
                    MessageBox.Show("Không có hình ảnh nào được chọn. Vui lòng chọn các hình ảnh và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (selectedShapes == null || selectedShapes.Count == 0)
                {
                    Logger.Error("Không có hình ảnh nào được chọn.");
                    MessageBox.Show("Không có hình ảnh nào được chọn.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Lấy tỷ lệ scale từ action panel
                double scalePercent = (double)_actionPanel.numScalePercent.Value / 100.0;

                int formattedCount = 0;
                foreach (Microsoft.Office.Interop.Excel.Shape shape in selectedShapes)
                {
                    try
                    {
                        // 1. Scale theo numScalePercent
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                        shape.Height = (float)(shape.Height * scalePercent);

                        // 2. Apply picture style
                        ApplyPictureStyleToShape(shape);

                        formattedCount++;
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Lỗi khi format hình ảnh '{shape.Name}': {ex.Message}");
                    }
                }

                Logger.Info($"Đã format {formattedCount} hình ảnh thành công.");
            }
            catch (Exception ex)
            {
                Logger.Error($"Có lỗi xảy ra khi format hình ảnh: {ex.Message}", ex);
                MessageBox.Show($"Có lỗi xảy ra khi format hình ảnh: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Xử lý sự kiện chèn nhiều hình ảnh
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                    Logger.Error("Không có workbook nào đang mở. Vui lòng mở một workbook và thử lại.");
                    MessageBox.Show("Không có workbook nào đang mở. Vui lòng mở một workbook và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                if (activeSheet == null)
                {
                    Logger.Error("Không có sheet nào đang được chọn. Vui lòng chọn một sheet và thử lại.");
                    MessageBox.Show("Không có sheet nào đang được chọn. Vui lòng chọn một sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Kiểm tra cell đang chọn
                Range activeCell = null;
                try { activeCell = app.ActiveCell as Range; } catch { }
                if (activeCell == null)
                {
                    Logger.Error("Không có ô nào đang được chọn hoặc lựa chọn không hợp lệ. Vui lòng chọn một ô và thử lại.");
                    MessageBox.Show("Không có ô nào đang được chọn hoặc lựa chọn không hợp lệ. Vui lòng chọn một ô và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Kiểm tra sheet có bị bảo vệ không
                if (activeSheet.ProtectContents || activeSheet.ProtectDrawingObjects || activeSheet.ProtectScenarios)
                {
                    Logger.Error("Sheet đang được bảo vệ. Vui lòng bỏ bảo vệ sheet và thử lại.");
                    MessageBox.Show("Sheet đang được bảo vệ. Vui lòng bỏ bảo vệ sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Đường dẫn thư mục chứa hình ảnh
                string folderPath = _actionPanel.txtImagePath.Text.Trim();

                // Tạo thư mục nếu chưa tồn tại
                if (!System.IO.Directory.Exists(folderPath))
                {
                    System.IO.Directory.CreateDirectory(folderPath);
                    Logger.Info($"Đã tạo thư mục '{folderPath}'. Vui lòng thêm hình ảnh vào thư mục này và thử lại.");
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
                    Logger.Warning($"Không tìm thấy file hình ảnh nào trong thư mục '{folderPath}'. Các định dạng được hỗ trợ: JPG, JPEG, PNG, BMP, GIF, TIFF");
                    MessageBox.Show($"Không tìm thấy file hình ảnh nào trong thư mục '{folderPath}'.\nCác định dạng được hỗ trợ: JPG, JPEG, PNG, BMP, GIF, TIFF", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Khởi tạo các biến cần thiết
                double resizeRate = (double)(_actionPanel.numScalePercent.Value / 100);
                bool insertOnNewPage = _actionPanel.chkInsertOnNewPage.Checked;
                Logger.Info($"Insert on new page mode: {insertOnNewPage}");

                // Tính toán vị trí bắt đầu
                var position = CalculateInitialImagePosition(activeSheet, activeCell, insertOnNewPage);
                int insertedCount = 0;
                int errorCount = 0;
                double maxBottomPosition = position.Top;
                double currentPageStartTop = position.Top; // Theo dõi vị trí bắt đầu của trang hiện tại

                // Chèn từng hình ảnh
                foreach (string imagePath in imageFiles)
                {
                    try
                    {
                        // Bước 1: Tải hình ảnh tạm thời để lấy chiều cao sau khi resize
                        var tempShape = activeSheet.Shapes.AddPicture(
                            imagePath,
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue,
                            0, 0, -1, -1
                        );
                        tempShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                        tempShape.Height = (float)(tempShape.Height * resizeRate);
                        double imageHeight = tempShape.Height;
                        tempShape.Delete(); // Xóa hình tạm

                        // Bước 2: Kiểm tra và thêm page break nếu cần (dựa vào chiều cao hình sắp chèn)
                        if (insertOnNewPage && insertedCount > 0)
                        {
                            AddPageBreakForNextImage(activeSheet, activeCell, ref position, ref currentPageStartTop, maxBottomPosition, insertedCount + 1, imageHeight);
                        }

                        // Bước 3: Chèn hình ảnh vào vị trí đã tính toán
                        var shape = activeSheet.Shapes.AddPicture(
                            imagePath,
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue,
                            (float)position.Left,
                            (float)position.Top,
                            -1, // Width - tự động
                            -1  // Height - tự động
                        );

                        // Điều chỉnh kích thước hình ảnh
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                        shape.Height = (float)(shape.Height * resizeRate);

                        // Áp dụng picture style
                        ApplyPictureStyleToShape(shape);

                        // Cập nhật vị trí cho hình ảnh tiếp theo
                        UpdatePositionForNextImage(activeSheet, activeCell, ref position, shape, insertOnNewPage);

                        // Theo dõi vị trí thấp nhất
                        maxBottomPosition = Math.Max(maxBottomPosition, shape.Top + shape.Height);

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
                            Logger.Warning($"Không thể xóa file {imagePath}: {deleteEx.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        Logger.Error($"Lỗi khi chèn hình ảnh {imagePath}: {ex.Message}");
                    }
                }

                // Tự động điều chỉnh print area sau khi chèn hình ảnh
                if (insertedCount > 0)
                {
                    try
                    {
                        AdjustPrintAreaForImages(activeSheet, activeCell, maxBottomPosition);
                        Logger.Info($"Print area adjusted for {insertedCount} inserted images");
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Failed to adjust print area: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Có lỗi xảy ra khi chèn hình ảnh: {ex.Message}", ex);
                MessageBox.Show($"Có lỗi xảy ra khi chèn hình ảnh: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Tính toán vị trí ban đầu để chèn hình ảnh
        /// </summary>
        private (double Top, double Left) CalculateInitialImagePosition(Worksheet activeSheet, Range activeCell, bool insertOnNewPage)
        {
            // Luôn chèn hình ảnh đầu tiên tại chính cell đang chọn
            double topLocation = (double)activeCell.Top;
            double leftLocation = (double)activeCell.Left;

            Logger.Debug($"Insert first image at cell {activeCell.Address[false, false]} (Row: {activeCell.Row}, Column: {activeCell.Column})");

            return (topLocation, leftLocation);
        }

        /// <summary>
        /// Thêm page break cho hình ảnh tiếp theo nếu cần thiết
        /// </summary>
        private void AddPageBreakForNextImage(Worksheet activeSheet, Range activeCell, ref (double Top, double Left) position, ref double currentPageStartTop, double maxBottomPosition, int imageIndex, double nextImageHeight)
        {
            try
            {
                // Tính toán chiều cao của trang hiện tại
                double pageHeight = activeCell.Height * PRINT_AREA_LAST_ROW_IDX;

                // Tính toán chiều cao đã sử dụng trên trang hiện tại
                double usedHeightOnCurrentPage = maxBottomPosition - currentPageStartTop;

                // Tính toán chiều cao sẽ sử dụng nếu chèn hình tiếp theo (bao gồm khoảng cách giữa các hình)
                double nextImageHeightWithSpacing = nextImageHeight + activeCell.Height;

                // Kiểm tra xem có cần tạo page break không
                // Tạo page break nếu: chiều cao đã dùng + chiều cao hình tiếp theo > chiều cao trang hiện tại
                if (usedHeightOnCurrentPage + nextImageHeightWithSpacing > pageHeight)
                {
                    // Chèn horizontal page break
                    int rowForPageBreak = Math.Max(1, (int)Math.Ceiling(maxBottomPosition / activeCell.Height) + 2);
                    Range breakRange = activeSheet.Cells[rowForPageBreak, 1];
                    activeSheet.HPageBreaks.Add(breakRange);

                    // Cập nhật vị trí cho ảnh mới tại row 2 của trang mới (cùng cột với cell ban đầu)
                    Range row2Cell = activeSheet.Cells[rowForPageBreak + 1, activeCell.Column];
                    position.Top = (double)row2Cell.Top;
                    position.Left = (double)row2Cell.Left;

                    // Cập nhật vị trí bắt đầu của trang mới
                    currentPageStartTop = position.Top;

                    Logger.Debug($"Added page break at row {rowForPageBreak}, image {imageIndex} will be placed at row {rowForPageBreak + 1} (row 2 of new page). Used height: {usedHeightOnCurrentPage:F1}, Next image height: {nextImageHeight:F1}, Total would be: {usedHeightOnCurrentPage + nextImageHeightWithSpacing:F1}, Page height: {pageHeight:F1}");
                }
                else
                {
                    Logger.Debug($"No page break needed for image {imageIndex}. Used height: {usedHeightOnCurrentPage:F1}, Next image height: {nextImageHeight:F1}, Total: {usedHeightOnCurrentPage + nextImageHeightWithSpacing:F1} / Page height: {pageHeight:F1}. Images will fit on current page.");
                }
            }
            catch (Exception pageBreakEx)
            {
                Logger.Warning($"Failed to add page break: {pageBreakEx.Message}");
                // Tiếp tục chèn ảnh mà không có page break
            }
        }

        /// <summary>
        /// Cập nhật vị trí cho hình ảnh tiếp theo
        /// </summary>
        private void UpdatePositionForNextImage(Worksheet activeSheet, Range activeCell, ref (double Top, double Left) position, Microsoft.Office.Interop.Excel.Shape shape, bool insertOnNewPage)
        {
            // Luôn cập nhật vị trí cho hình ảnh tiếp theo (xếp theo chiều dọc)
            // Page break sẽ được xử lý riêng trong AddPageBreakForNextImage
            position.Top += shape.Height + activeCell.Height;

            Logger.Debug($"Updated position for next image: Top={position.Top:F1}, Left={position.Left:F1}");
        }

        /// <summary>
        /// ThisAddIn_Startup
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Logger.Info("Excel Custom Add-in starting up...");

            // 4939, 5073, 5175
            try
            {
                // Khởi tạo cấu hình từ XML
                SheetConfigManager.LoadConfiguration();
                Logger.Info("Sheet configuration loaded successfully");

                // Log thông tin về cấu hình logging
                var loggingConfig = SheetConfigManager.GetLoggingConfig();
                Logger.Info($"Logger configured - Directory: {(string.IsNullOrEmpty(loggingConfig.LogDirectory) ? "Default C:\\ExcelCustomAddin" : loggingConfig.LogDirectory)}, File: {loggingConfig.LogFileName}, Debug: {loggingConfig.EnableDebugOutput}");
                Logger.Info($"Log file path: {Logger.GetLogFilePath()}");

                // Thiết lập debug logging dựa trên cấu hình
                var generalConfig = SheetConfigManager.GetGeneralConfig();
                if (generalConfig != null)
                {
                    Logger.SetDebugEnabled(generalConfig.EnableDebugLog);
                    Logger.Debug($"Debug logging {(generalConfig.EnableDebugLog ? "enabled" : "disabled")}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error loading sheet configuration", ex);
            }

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
            Logger.Debug($"Application_NewWorkbook called for: {Wb?.Name}");
            this.LoadTemplate(Wb);
            this.CreateActionsPane(Wb);
        }

        /// <summary>
        /// Application_WorkbookActivate
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookOpen(Workbook Wb)
        {
            Logger.Debug($"Application_WorkbookOpen called for: {Wb?.Name}");
            this.LoadTemplate(Wb);
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
                Logger.Debug($"File path refreshed after save for: {Wb.Name}");
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
                    Logger.Debug($"CreateActionsPane called for: {workbookKey}");

                    // Kiểm tra xem action panel đã được tạo cho workbook này chưa
                    if (CreatedActionPanes.Contains(workbookKey))
                    {
                        Logger.Debug($"Action panel already exists for: {workbookKey}, updating only");
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

                    Logger.Debug($"Creating new action panel for: {workbookKey}");

                    // Get Active ActionsPanel
                    myCustomTaskPane = TaskPaneManager.GetTaskPane(Wb.Name, "WORKSHEET TOOLS", () => new ActionPanelControl());
                    _actionPanel = (ActionPanelControl)myCustomTaskPane?.Control;

                    if (_actionPanel != null)
                    {
                        // Hủy đăng ký các event cũ trước khi đăng ký mới để tránh đăng ký trùng lặp
                        _actionPanel.CreateEvidenceEvent -= this.CreateEvidence;
                        _actionPanel.FormatImagesEvent -= this.FormatImages;
                        _actionPanel.FormatDocumentEvent -= this.FormatDocument;
                        _actionPanel.ChangeSheetNameEvent -= this.ChangeSheetName;
                        _actionPanel.InsertMultipleImagesEvent -= this.InsertMultipleImages;
                        _actionPanel.PinSheetEvent -= this.PinSheet;
                        _actionPanel.listofSheet.SelectedIndexChanged -= this.ListOfSheet_SelectionChanged;

                        // Đăng ký các event mới
                        _actionPanel.CreateEvidenceEvent += this.CreateEvidence;
                        _actionPanel.FormatImagesEvent += this.FormatImages;
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
                        Logger.Debug($"Action panel created and marked for: {workbookKey}");
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
                    Logger.Error("Không có workbook nào đang mở trong RenameWorksheet");
                    MessageBox.Show("Không có workbook nào đang mở. Vui lòng mở một workbook và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Lấy tên sheet từ ListView thay vì active sheet
                if (_actionPanel.listofSheet.SelectedItems.Count == 0)
                {
                    Logger.Error("Không có sheet nào được chọn từ danh sách để đổi tên");
                    MessageBox.Show("Vui lòng chọn một sheet từ danh sách để đổi tên.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var selectedItem = _actionPanel.listofSheet.SelectedItems[0].Tag as SheetInfo;
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
                if (_actionPanel != null)
                {
                    _actionPanel.BindSheetList(this.GetListOfSheet(), newSheetName);
                }

                Logger.Info($"Successfully renamed sheet from '{oldSheetName}' to '{newSheetName}', updated {updatedLinksCount} hyperlinks");
                MessageBox.Show($"Đã đổi tên sheet từ '{oldSheetName}' thành '{newSheetName}' thành công.\nĐã cập nhật {updatedLinksCount} hyperlinks.",
                    "Hoàn thành", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Logger.Error($"Có lỗi xảy ra khi đổi tên sheet: {ex.Message}", ex);
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

        /// <summary>
        /// FormatDocument
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormatDocument(object sender, EventArgs e)
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

                // Validate inputs
                Range selectedRange;
                if (!ValidateEvidenceCreationInputs(activeWorkbook, activeSheet, app, out selectedRange))
                {
                    return;
                }

                // Get cells to process (handles both single and multiple cells, merged cells)
                var cellsToProcess = GetCellsToProcess(selectedRange);
                bool isMultipleCells = selectedRange.Cells.Count > 1;

                // Process each cell and create evidence sheets
                var (createdSheets, existingSheets, errorMessages) = ProcessEvidenceCells(
                    cellsToProcess, activeWorkbook, activeSheet);

                // Show results and update UI
                HandleEvidenceCreationResults(
                    isMultipleCells, createdSheets, existingSheets, errorMessages,
                    activeWorkbook, activeSheet);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi xảy ra khi tạo sheet bằng chứng: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Validate inputs for evidence creation
        /// </summary>
        private bool ValidateEvidenceCreationInputs(Workbook activeWorkbook, Worksheet activeSheet, Microsoft.Office.Interop.Excel.Application app, out Range selectedRange)
        {
            selectedRange = null;

            if (activeWorkbook == null)
            {
                Logger.Error("Không có workbook nào đang mở trong CreateEvidence");
                MessageBox.Show("Không có workbook nào đang mở. Vui lòng mở một workbook và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (activeSheet == null)
            {
                Logger.Error("Không có sheet nào đang được chọn trong CreateEvidence");
                MessageBox.Show("Không có sheet nào đang được chọn. Vui lòng chọn một sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            try { selectedRange = app.Selection as Range; } catch { }
            if (selectedRange == null)
            {
                MessageBox.Show("Không có ô nào đang được chọn hoặc lựa chọn không hợp lệ. Vui lòng chọn một ô và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (activeSheet.ProtectContents || activeSheet.ProtectDrawingObjects || activeSheet.ProtectScenarios)
            {
                MessageBox.Show("Sheet đang được bảo vệ. Vui lòng bỏ bảo vệ sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Get cells to process from selected range (handles merged cells)
        /// </summary>
        private List<Range> GetCellsToProcess(Range selectedRange)
        {
            var cellsToProcess = new List<Range>();
            var processedMergedAreas = new HashSet<string>();
            bool isMultipleCells = selectedRange.Cells.Count > 1;

            if (isMultipleCells)
            {
                foreach (Range cell in selectedRange.Cells)
                {
                    if (cell.MergeCells)
                    {
                        Range mergedArea = cell.MergeArea;
                        string mergedAreaAddress = mergedArea.Address[true, true];

                        if (!processedMergedAreas.Contains(mergedAreaAddress))
                        {
                            Range firstCell = mergedArea.Cells[1, 1];
                            cellsToProcess.Add(firstCell);
                            processedMergedAreas.Add(mergedAreaAddress);
                            Logger.Debug($"Merged cell detected at {cell.Address[false, false]}, using first cell {firstCell.Address[false, false]}");
                        }
                    }
                    else
                    {
                        cellsToProcess.Add(cell);
                    }
                }
            }
            else
            {
                if (selectedRange.MergeCells)
                {
                    Range mergedArea = selectedRange.MergeArea;
                    Range firstCell = mergedArea.Cells[1, 1];
                    cellsToProcess.Add(firstCell);
                    Logger.Debug($"Single merged cell detected at {selectedRange.Address[false, false]}, using first cell {firstCell.Address[false, false]}");
                }
                else
                {
                    cellsToProcess.Add(selectedRange);
                }
            }

            return cellsToProcess;
        }

        /// <summary>
        /// Process evidence cells and create/link sheets
        /// </summary>
        private (List<string> createdSheets, List<string> existingSheets, List<string> errorMessages)
            ProcessEvidenceCells(List<Range> cellsToProcess, Workbook activeWorkbook, Worksheet activeSheet)
        {
            var createdSheets = new List<string>();
            var existingSheets = new List<string>();
            var errorMessages = new List<string>();

            // Build worksheet dictionary for faster lookup
            var worksheetDict = new Dictionary<string, Worksheet>(StringComparer.OrdinalIgnoreCase);
            foreach (Worksheet ws in activeWorkbook.Worksheets)
            {
                worksheetDict[ws.Name] = ws;
            }

            foreach (var cell in cellsToProcess)
            {
                try
                {
                    string cellValue = GetOrGenerateCellValue(cell, activeSheet, cellsToProcess.Count);

                    if (string.IsNullOrEmpty(cellValue))
                    {
                        errorMessages.Add($"Ô {cell.Address[false, false]} đang để trống. Trường hợp chọn nhiều ô thì cần nhập sẵn giá trị cho các ô đã chọn");
                        continue;
                    }

                    if (worksheetDict.TryGetValue(cellValue, out Worksheet existingSheet))
                    {
                        CreateHyperlinkToExistingSheet(cell, activeSheet, existingSheet, cellValue);
                        existingSheets.Add(cellValue);
                    }
                    else
                    {
                        Worksheet newWs = CreateNewEvidenceSheet(activeWorkbook, cellValue);
                        CreateHyperlinkToNewSheet(cell, activeSheet, newWs, cellValue);
                        createdSheets.Add(cellValue);
                        worksheetDict[cellValue] = newWs; // Add to cache
                    }
                }
                catch (Exception ex)
                {
                    errorMessages.Add($"Lỗi khi xử lý ô {cell.Address[false, false]}: {ex.Message}");
                }
            }

            return (createdSheets, existingSheets, errorMessages);
        }

        /// <summary>
        /// Get cell value or generate auto sheet name if empty
        /// </summary>
        private string GetOrGenerateCellValue(Range cell, Worksheet activeSheet, int cellsToProcessCount)
        {
            string cellValue = cell.Value2 != null ? cell.Value2.ToString().Trim() : "";

            if (string.IsNullOrEmpty(cellValue) && cellsToProcessCount <= 1)
            {
                string currentSheetName = activeSheet.Name;
                if (currentSheetName == "共通" || currentSheetName == "テスト項目")
                {
                    cellValue = GenerateAutoSheetName(activeSheet, cell.Column, currentSheetName);
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        cell.Value2 = cellValue;
                        Logger.Debug($"Auto-generated sheet name '{cellValue}' for cell {cell.Address[false, false]} (Column: {cell.Column})");
                    }
                }
            }

            return cellValue;
        }

        /// <summary>
        /// Create hyperlink to existing sheet and setup back button with named range reference
        /// </summary>
        private void CreateHyperlinkToExistingSheet(Range cell, Worksheet sourceSheet, Worksheet existingSheet, string sheetName)
        {
            // Tạo hyperlink đến sheet đã tồn tại
            sourceSheet.Hyperlinks.Add(cell, "", $"'{sheetName}'!A1", Type.Missing, sheetName);
            cell.Font.Name = EVIDENCE_FONT_NAME;

            // Tạo hoặc lấy named range cho cell gốc
            string namedRangeName = GetOrCreateNamedRangeForCell(cell, sourceSheet);

            // Setup back button
            int aColumnIndex = GetColumnIndex("A");
            Range backCell = existingSheet.Cells[1, aColumnIndex];
            backCell.Value2 = "<Back";

            try
            {
                if (backCell.Hyperlinks.Count > 0)
                {
                    backCell.Hyperlinks.Delete();
                }
            }
            catch { }

            // Tạo hyperlink back về cell gốc - sử dụng named range nếu có, ngược lại dùng địa chỉ trực tiếp
            string backAddress;
            if (!string.IsNullOrEmpty(namedRangeName))
            {
                backAddress = namedRangeName; // Sử dụng named range
                Logger.Debug($"Back button (existing sheet) sẽ reference đến named range: {namedRangeName}");
            }
            else
            {
                backAddress = $"'{sourceSheet.Name}'!{cell.Address[false, false]}"; // Fallback về địa chỉ trực tiếp
                Logger.Debug($"Back button (existing sheet) sẽ reference đến địa chỉ trực tiếp: {backAddress}");
            }

            existingSheet.Hyperlinks.Add(backCell, "", backAddress, Type.Missing, "<Back");
            var config = SheetConfigManager.GetGeneralConfig();
            backCell.Font.Name = config.BackButtonFontName;
        }

        /// <summary>
        /// Create hyperlink to new sheet and setup back button with named range reference
        /// </summary>
        private void CreateHyperlinkToNewSheet(Range cell, Worksheet sourceSheet, Worksheet newSheet, string sheetName)
        {
            // Tạo hyperlink đến sheet mới
            sourceSheet.Hyperlinks.Add(cell, "", $"'{sheetName}'!A1", Type.Missing, sheetName);
            cell.Font.Name = EVIDENCE_FONT_NAME;

            // Tạo named range cho cell gốc trước khi setup back button
            string namedRangeName = null;
            try
            {
                // Lưu trữ active cell hiện tại
                var originalActiveCell = Globals.ThisAddIn.Application.ActiveCell as Range;

                // Activate source sheet và cell để tạo named range
                sourceSheet.Activate();
                cell.Select();

                // Tạo named range cho cell gốc và lấy tên
                namedRangeName = CreateNamedRange();

                // Khôi phục active cell gốc nếu có
                if (originalActiveCell != null)
                {
                    try
                    {
                        originalActiveCell.Worksheet.Activate();
                        originalActiveCell.Select();
                    }
                    catch { /* Ignore errors when restoring */ }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Không thể tạo named range cho back button: {ex.Message}");
            }

            // Setup back button
            int aColumnIndex = GetColumnIndex("A");
            Range backCell = newSheet.Cells[1, aColumnIndex];
            backCell.Value2 = "<Back";

            // Tạo hyperlink back về cell gốc - sử dụng named range nếu có, ngược lại dùng địa chỉ trực tiếp
            string backAddress;
            if (!string.IsNullOrEmpty(namedRangeName))
            {
                backAddress = namedRangeName; // Sử dụng named range
                Logger.Debug($"Back button sẽ reference đến named range: {namedRangeName}");
            }
            else
            {
                backAddress = $"'{sourceSheet.Name}'!{cell.Address[false, false]}"; // Fallback về địa chỉ trực tiếp
                Logger.Debug($"Back button sẽ reference đến địa chỉ trực tiếp: {backAddress}");
            }

            newSheet.Hyperlinks.Add(backCell, "", backAddress, Type.Missing, "<Back");
            var config = SheetConfigManager.GetGeneralConfig();
            backCell.Font.Name = config.BackButtonFontName;
        }

        /// <summary>
        /// Handle evidence creation results (show messages and update UI)
        /// </summary>
        private void HandleEvidenceCreationResults(
            bool isMultipleCells,
            List<string> createdSheets,
            List<string> existingSheets,
            List<string> errorMessages,
            Workbook activeWorkbook,
            Worksheet activeSheet)
        {
            if (isMultipleCells)
            {
                ShowMultipleCellsResults(existingSheets, errorMessages);
                UpdateActionPanel(activeSheet.Name);
            }
            else
            {
                HandleSingleCellResult(createdSheets, activeWorkbook, activeSheet);
            }
        }

        /// <summary>
        /// Show results for multiple cells processing
        /// </summary>
        private void ShowMultipleCellsResults(List<string> existingSheets, List<string> errorMessages)
        {
            if (existingSheets.Count == 0 && errorMessages.Count == 0)
            {
                return;
            }

            var resultMessage = new StringBuilder();

            if (existingSheets.Count > 0)
            {
                resultMessage.AppendLine($"\n⚠ {existingSheets.Count} sheet(s) đã tồn tại (chỉ tạo hyperlink):");
                foreach (var sheetName in existingSheets)
                {
                    resultMessage.AppendLine($"  • {sheetName}");
                }
            }

            if (errorMessages.Count > 0)
            {
                resultMessage.AppendLine($"\n❌ {errorMessages.Count} lỗi:");
                foreach (var error in errorMessages)
                {
                    resultMessage.AppendLine($"  • {error}");
                }
            }

            if (resultMessage.Length > 0)
            {
                MessageBox.Show(resultMessage.ToString(), "Kết quả tạo Evidence Sheets", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        /// <summary>
        /// Handle result for single cell processing (focus on new sheet if created)
        /// </summary>
        private void HandleSingleCellResult(List<string> createdSheets, Workbook activeWorkbook, Worksheet activeSheet)
        {
            if (createdSheets.Count > 0)
            {
                string newSheetName = createdSheets[0];
                foreach (Worksheet ws in activeWorkbook.Worksheets)
                {
                    if (ws.Name == newSheetName)
                    {
                        ws.Activate();
                        break;
                    }
                }
                UpdateActionPanel(newSheetName);
            }
            else
            {
                UpdateActionPanel(activeSheet.Name);
            }
        }

        /// <summary>
        /// Update action panel with current sheet
        /// </summary>
        private void UpdateActionPanel(string currentSheetName)
        {
            if (_actionPanel != null)
            {
                _actionPanel.BindSheetList(this.GetListOfSheet(), currentSheetName);
            }
        }

        /// <summary>
        /// Tạo một Evidence Sheet mới với tất cả cấu hình cần thiết
        /// </summary>
        /// <param name="activeWorkbook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private Worksheet CreateNewEvidenceSheet(Workbook activeWorkbook, string sheetName)
        {
            var config = SheetConfigManager.GetGeneralConfig();

            // Tạo sheet mới
            Worksheet newWs = activeWorkbook.Worksheets.Add(Type.Missing, activeWorkbook.Worksheets[activeWorkbook.Worksheets.Count]);
            newWs.Name = sheetName;

            newWs.Columns.ColumnWidth = config.ColumnWidth;
            newWs.Rows.RowHeight = config.RowHeight;

            // Thiết lập font chữ cho toàn bộ sheet
            newWs.Cells.Font.Name = EVIDENCE_FONT_NAME;
            newWs.Cells.Font.Size = config.FontSize;

            // Đặt giá trị vào ô AR1 để mở rộng used range
            int azColumnIndex = GetColumnIndex(PAGE_BREAK_COLUMN_NAME);
            newWs.Cells[1, azColumnIndex].Value2 = " ";
            newWs.PageSetup.Orientation = config.PageOrientation == "Landscape" ? XlPageOrientation.xlLandscape : XlPageOrientation.xlPortrait;
            newWs.PageSetup.PaperSize = config.PaperSize == "A4" ? XlPaperSize.xlPaperA4 : XlPaperSize.xlPaperA3; // Add more cases if needed
            newWs.PageSetup.PrintArea = "$A$1:$" + PAGE_BREAK_COLUMN_NAME + "$" + PRINT_AREA_LAST_ROW_IDX.ToString();
            newWs.PageSetup.Zoom = config.Zoom;
            newWs.PageSetup.FitToPagesWide = config.FitToPagesWide;
            newWs.PageSetup.FitToPagesTall = config.FitToPagesTall;

            // Thiết lập lề trang tối ưu cho Windows (đơn vị: inches)
            newWs.PageSetup.LeftMargin = newWs.Application.InchesToPoints(config.LeftMargin);   // Lề trái
            newWs.PageSetup.RightMargin = newWs.Application.InchesToPoints(config.RightMargin);  // Lề phải
            newWs.PageSetup.TopMargin = newWs.Application.InchesToPoints(config.TopMargin);       // Lề trên
            newWs.PageSetup.BottomMargin = newWs.Application.InchesToPoints(config.BottomMargin); // Lề dưới
            newWs.PageSetup.HeaderMargin = newWs.Application.InchesToPoints(config.HeaderMargin);  // Lề header
            newWs.PageSetup.FooterMargin = newWs.Application.InchesToPoints(config.FooterMargin);  // Lề footer

            // Thiết lập center on page theo chiều horizontal
            newWs.PageSetup.CenterHorizontally = config.CenterHorizontally;

            // Thiết lập view mode thành Page Break Preview
            try
            {
                var window = newWs.Application.ActiveWindow;
                if (window != null)
                {
                    window.View = XlWindowView.xlPageBreakPreview;
                    // Thiết lập zoom về 100%
                    window.Zoom = config.WindowZoom;
                }
            }
            catch (Exception viewEx)
            {
                // Log error nhưng không làm gián đoạn quá trình tạo sheet
                Logger.Warning($"Error setting page break preview or zoom: {viewEx.Message}");
            }

            return newWs;
        }

        /// <summary>
        /// Thiết lập nút Back và hyperlink trong sheet (mới hoặc đã tồn tại)
        /// </summary>
        /// <param name="targetSheet">Sheet đích (mới hoặc đã tồn tại)</param>
        /// <param name="sourceSheet">Sheet gốc chứa cell đã click</param>
        /// <param name="sourceCell">Cell gốc đã được click</param>
        private void SetupBackButtonAndHyperlink(Worksheet targetSheet, Worksheet sourceSheet, Range sourceCell)
        {
            try
            {
                // Đặt giá trị "<Back" (Back) vào ô A1 của sheet đích
                int aColumnIndex = GetColumnIndex("A"); // A = 1
                Range backCell = targetSheet.Cells[1, aColumnIndex];

                // Xóa hyperlink cũ nếu có
                try
                {
                    if (backCell.Hyperlinks.Count > 0)
                    {
                        backCell.Hyperlinks.Delete();
                    }
                }
                catch (Exception ex)
                {
                    Logger.Debug($"No existing hyperlink to delete in {targetSheet.Name}: {ex.Message}");
                }

                // Cập nhật giá trị và định dạng cell
                backCell.Value2 = "<Back";

                // Định dạng cell chứa nút Back
                backCell.Font.Name = EVIDENCE_FONT_NAME;
                backCell.Font.Size = 12;
                backCell.Font.Bold = true;
                backCell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                backCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                backCell.VerticalAlignment = XlVAlign.xlVAlignCenter;

                // Thêm border cho cell Back
                backCell.Borders.LineStyle = XlLineStyle.xlContinuous;
                backCell.Borders.Weight = XlBorderWeight.xlMedium;
                backCell.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

                // Đặt background color nhẹ cho cell Back
                backCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightBlue);

                // Tạo hyperlink "Back" từ ô A1 của sheet đích về ô gốc
                string sourceAddress = $"'{sourceSheet.Name}'!{sourceCell.Address[false, false]}";
                targetSheet.Hyperlinks.Add(backCell, "", sourceAddress, Type.Missing, "<Back");

                // Đảm bảo font cho Back hyperlink cell (đã set ở trên nhưng đảm bảo)
                backCell.Font.Name = EVIDENCE_FONT_NAME;

                // Cập nhật hoặc thêm comment/note cho cell Back
                try
                {
                    // Xóa comment cũ nếu có
                    if (backCell.Comment != null)
                    {
                        backCell.Comment.Delete();
                    }

                    // Thêm comment mới
                    backCell.AddComment($"Click để quay về ô {sourceCell.Address[false, false]} trong sheet '{sourceSheet.Name}'");
                    backCell.Comment.Shape.TextFrame.AutoSize = true;
                }
                catch (Exception commentEx)
                {
                    Logger.Warning($"Error adding comment to back button: {commentEx.Message}");
                }

                // Cập nhật title cho sheet ở ô B1 (chỉ nếu chưa có hoặc cần cập nhật)
                int bColumnIndex = GetColumnIndex("B"); // B = 2
                Range titleCell = targetSheet.Cells[1, bColumnIndex];
                if (titleCell.Value2 == null || string.IsNullOrEmpty(titleCell.Value2.ToString()) ||
                    !titleCell.Value2.ToString().StartsWith("Evidence:"))
                {
                    titleCell.Value2 = $"Evidence: {targetSheet.Name}";
                    titleCell.Font.Name = EVIDENCE_FONT_NAME;
                    titleCell.Font.Size = 14;
                    titleCell.Font.Bold = true;
                    titleCell.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    titleCell.VerticalAlignment = XlVAlign.xlVAlignCenter;
                }

                Logger.Debug($"Back button and hyperlink setup completed for sheet '{targetSheet.Name}', linking to '{sourceSheet.Name}'!{sourceCell.Address[false, false]}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error setting up back button and hyperlink: {ex.Message}", ex);
                // Fallback: tạo hyperlink đơn giản nếu có lỗi
                try
                {
                    int aColumnIndex = GetColumnIndex("A"); // A = 1
                    Range fallbackCell = targetSheet.Cells[1, aColumnIndex];
                    // Xóa hyperlink cũ nếu có
                    if (fallbackCell.Hyperlinks.Count > 0)
                    {
                        fallbackCell.Hyperlinks.Delete();
                    }

                    fallbackCell.Value2 = "<Back";
                    string sourceAddress = $"'{sourceSheet.Name}'!{sourceCell.Address[false, false]}";
                    targetSheet.Hyperlinks.Add(fallbackCell, "", sourceAddress, Type.Missing, "<Back");

                    // Thiết lập font cho fallback Back hyperlink cell
                    fallbackCell.Font.Name = EVIDENCE_FONT_NAME;

                    Logger.Info("Fallback back button created successfully");
                }
                catch (Exception fallbackEx)
                {
                    Logger.Error($"Error creating fallback back button: {fallbackEx.Message}", fallbackEx);
                }
            }
        }

        /// <summary>
        /// Tự động điều chỉnh print area để bao gồm các hình ảnh đã chèn
        /// </summary>
        /// <param name="worksheet">Worksheet chứa hình ảnh</param>
        /// <param name="startCell">Cell bắt đầu chèn hình ảnh</param>
        /// <param name="maxBottomPosition">Vị trí thấp nhất của hình ảnh</param>
        private void AdjustPrintAreaForImages(Worksheet worksheet, Range startCell, double maxBottomPosition)
        {
            try
            {
                var config = SheetConfigManager.GetGeneralConfig();
                Logger.Debug($"Adjusting print area for images in sheet: {worksheet.Name}");

                // Tìm tất cả shapes trong worksheet để tính toán vùng in chính xác
                double minLeft = double.MaxValue;
                double maxRight = double.MinValue;
                double minTop = double.MaxValue;
                double maxBottom = double.MinValue;

                bool hasShapes = false;
                int imageCount = 0;
                foreach (Microsoft.Office.Interop.Excel.Shape shape in worksheet.Shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                    {
                        hasShapes = true;
                        imageCount++;
                        double shapeLeft = shape.Left;
                        double shapeTop = shape.Top;
                        double shapeRight = shape.Left + shape.Width;
                        double shapeBottom = shape.Top + shape.Height;

                        minLeft = Math.Min(minLeft, shapeLeft);
                        maxRight = Math.Max(maxRight, shapeRight);
                        minTop = Math.Min(minTop, shapeTop);
                        maxBottom = Math.Max(maxBottom, shapeBottom);

                        Logger.Debug($"Image {imageCount}: Left={shapeLeft:F1}, Top={shapeTop:F1}, Right={shapeRight:F1}, Bottom={shapeBottom:F1}");
                    }
                }

                if (!hasShapes)
                {
                    Logger.Debug("No images found in worksheet, skipping print area adjustment");
                    return;
                }

                Logger.Debug($"Found {imageCount} images. Content bounds: Left={minLeft:F1}, Top={minTop:F1}, Right={maxRight:F1}, Bottom={maxBottom:F1}");

                // Chuyển đổi từ tọa độ pixel sang cell addresses bằng cách tính toán trực tiếp
                int startColumn = 1;
                int startRow = 1;
                int endColumn = 1;
                int endRow = 1;

                int aColumnIndex = GetColumnIndex("A");

                try
                {
                    // Tìm cell chứa góc trên-trái của vùng hình ảnh
                    Range topLeftCell = null;
                    Range bottomRightCell = null;

                    // Sử dụng ActiveWindow để convert từ pixel sang cell
                    if (worksheet.Application.ActiveWindow != null)
                    {
                        try
                        {
                            topLeftCell = worksheet.Application.ActiveWindow.RangeFromPoint((int)minLeft, (int)minTop);
                            bottomRightCell = worksheet.Application.ActiveWindow.RangeFromPoint((int)maxRight, (int)maxBottom);
                        }
                        catch (Exception ex)
                        {
                            Logger.Warning($"RangeFromPoint failed: {ex.Message}");
                        }
                    }

                    // Fallback method: ước tính dựa trên kích thước cell tiêu chuẩn
                    if (topLeftCell == null || bottomRightCell == null)
                    {
                        Logger.Debug("Using fallback calculation method");

                        // Lấy kích thước cell thực tế từ worksheet
                        double actualCellWidth = (double)worksheet.Cells[1, aColumnIndex].Width;
                        double actualCellHeight = (double)worksheet.Cells[1, aColumnIndex].Height;

                        Logger.Debug($"Actual cell dimensions: Width={actualCellWidth:F1}, Height={actualCellHeight:F1}");

                        startColumn = Math.Max(1, (int)(minLeft / actualCellWidth) + 1);
                        startRow = Math.Max(1, (int)(minTop / actualCellHeight) + 1);
                        endColumn = Math.Min(worksheet.Columns.Count, (int)(maxRight / actualCellWidth) + 2);
                        endRow = Math.Min(worksheet.Rows.Count, (int)(maxBottom / actualCellHeight) + 3);

                        // Đảm bảo vùng in ít nhất có 5 dòng và 5 cột
                        endRow = Math.Max(endRow, startRow + 4);
                        endColumn = Math.Max(endColumn, startColumn + 4);

                        Logger.Debug($"Fallback calculation: StartCol={startColumn}, StartRow={startRow}, EndCol={endColumn}, EndRow={endRow}");
                    }
                    else
                    {
                        // Sử dụng kết quả từ RangeFromPoint
                        startColumn = Math.Max(1, topLeftCell.Column - 1);
                        startRow = Math.Max(1, topLeftCell.Row - 1);
                        endColumn = Math.Min(worksheet.Columns.Count, bottomRightCell.Column + 1);
                        endRow = Math.Min(worksheet.Rows.Count, bottomRightCell.Row + 2);

                        Logger.Debug($"RangeFromPoint calculation: TopLeft=({topLeftCell.Column},{topLeftCell.Row}), BottomRight=({bottomRightCell.Column},{bottomRightCell.Row})");
                        Logger.Debug($"RangeFromPoint result: StartCol={startColumn}, StartRow={startRow}, EndCol={endColumn}, EndRow={endRow}");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"Error calculating cell range: {ex.Message}");

                    // Sử dụng used range làm fallback cuối cùng
                    Range usedRange = worksheet.UsedRange;
                    if (usedRange != null)
                    {
                        startColumn = usedRange.Column;
                        startRow = usedRange.Row;
                        endColumn = usedRange.Column + usedRange.Columns.Count - 1;
                        endRow = usedRange.Row + usedRange.Rows.Count - 1;
                        Logger.Debug($"Used range fallback: StartCol={startColumn}, StartRow={startRow}, EndCol={endColumn}, EndRow={endRow}");
                    }
                }

                // Cố định vùng in từ A1 đến AR theo chiều cao của hình ảnh
                int azColumnIndex = GetColumnIndex(PAGE_BREAK_COLUMN_NAME);

                startColumn = aColumnIndex;
                startRow = 1;
                endColumn = azColumnIndex;

                // Tính toán endRow dựa trên vị trí thấp nhất của hình ảnh
                if (maxBottom > 0)
                {
                    // Lấy kích thước cell thực tế để tính toán dòng cuối
                    double actualCellHeight = (double)worksheet.Cells[1, aColumnIndex].Height;
                    endRow = Math.Max(PRINT_AREA_LAST_ROW_IDX, (int)(maxBottom / actualCellHeight) + 1); // Tối thiểu PRINT_AREA_LAST_ROW_IDX dòng, chỉ thêm 1 dòng buffer
                    Logger.Debug($"Calculated endRow based on image bottom: {endRow} (maxBottom={maxBottom:F1}, cellHeight={actualCellHeight:F1})");
                }
                else
                {
                    endRow = PRINT_AREA_LAST_ROW_IDX;
                }

                // Đảm bảo không vượt quá giới hạn worksheet
                endRow = Math.Min(endRow, worksheet.Rows.Count);

                Logger.Debug($"Fixed print area bounds: A1:AZ{endRow} (Column {startColumn}-{endColumn}, Row 1-{endRow})");

                // Tạo print area mới với định dạng cố định A1:AR{endRow}
                string newPrintArea = "$A$1:$" + PAGE_BREAK_COLUMN_NAME + "$" + endRow;
                worksheet.PageSetup.PrintArea = newPrintArea;

                Logger.Info($"Print area updated to: {newPrintArea} (covers {imageCount} images, fixed width A-AZ)");

                // Với print area cố định A1:AZ, luôn sử dụng Landscape orientation
                worksheet.PageSetup.Orientation = XlPageOrientation.xlLandscape;
                Logger.Debug("Set page orientation to Landscape (fixed for A1:AZ format)");

                // Thiết lập fit to page
                worksheet.PageSetup.Zoom = config.Zoom; // Thiết lập scaling từ config
                worksheet.PageSetup.FitToPagesWide = config.FitToPagesWide; // Tắt FitToPagesWide
                worksheet.PageSetup.FitToPagesTall = config.FitToPagesTall; // Tắt FitToPagesTall

                // Thiết lập margins tối ưu cho hình ảnh với format A1:AZ
                worksheet.PageSetup.LeftMargin = worksheet.Application.InchesToPoints(config.LeftMargin);
                worksheet.PageSetup.RightMargin = worksheet.Application.InchesToPoints(config.RightMargin);
                worksheet.PageSetup.TopMargin = worksheet.Application.InchesToPoints(config.TopMargin);
                worksheet.PageSetup.BottomMargin = worksheet.Application.InchesToPoints(config.BottomMargin);

                // Thiết lập center on page theo chiều horizontal
                worksheet.PageSetup.CenterHorizontally = config.CenterHorizontally;

                // Thiết lập kích cỡ trang từ config
                worksheet.PageSetup.PaperSize = config.PaperSize == "A4" ? XlPaperSize.xlPaperA4 : XlPaperSize.xlPaperA3; // Add more cases if needed
                Logger.Debug($"Set paper size to {config.PaperSize} for A1:AZ print area");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error adjusting print area: {ex.Message}", ex);
                throw; // Re-throw để caller có thể xử lý
            }
        }

        /// <summary>
        /// Chuyển đổi chữ cái cột thành số (A=1, B=2, ..., Z=26, AA=27, AB=28, ..., AZ=52, BA=53, ...)
        /// </summary>
        /// <param name="columnLetter">Chữ cái cột (A, B, C, ..., AA, AB, ...)</param>
        /// <returns>Số cột (1-based)</returns>
        private int GetColumnIndex(string columnLetter)
        {
            if (string.IsNullOrEmpty(columnLetter))
            {
                throw new ArgumentException("Column letter cannot be null or empty", nameof(columnLetter));
            }

            columnLetter = columnLetter.ToUpper();
            int columnIndex = 0;

            for (int i = 0; i < columnLetter.Length; i++)
            {
                char c = columnLetter[i];
                if (c < 'A' || c > 'Z')
                {
                    throw new ArgumentException($"Invalid character '{c}' in column letter", nameof(columnLetter));
                }

                columnIndex = columnIndex * 26 + (c - 'A' + 1);
            }

            return columnIndex;
        }

        /// <summary>
        /// Reset print area về toàn bộ used range
        /// </summary>
        /// <param name="worksheet">Worksheet cần reset</param>
        public void ResetPrintArea(Worksheet worksheet)
        {
            try
            {
                worksheet.PageSetup.PrintArea = "";
                Logger.Info($"Print area reset for sheet: {worksheet.Name}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error resetting print area: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Tạo tên sheet tự động dựa trên pattern từ cấu hình XML
        /// </summary>
        /// <param name="worksheet">Worksheet hiện tại</param>
        /// <param name="columnIndex">Cột đang được chọn</param>
        /// <param name="sheetName">Tên sheet hiện tại</param>
        /// <returns>Tên sheet mới được tạo</returns>
        private string GenerateAutoSheetName(Worksheet worksheet, int columnIndex, string sheetName)
        {
            try
            {
                // Lấy cấu hình từ XML
                var sheetConfig = SheetConfigManager.GetSheetConfig(sheetName);
                if (sheetConfig == null)
                {
                    Logger.Debug($"No configuration found for sheet: {sheetName}");
                    return ""; // Không có cấu hình cho sheet này
                }

                string prefix = sheetConfig.Prefix;
                string numberFormat = sheetConfig.NumberFormat;
                var generalConfig = SheetConfigManager.GetGeneralConfig();

                // Tìm giá trị lớn nhất trong cột
                int maxNumber = generalConfig.StartingNumber - 1; // Bắt đầu từ giá trị cấu hình - 1

                // Tìm cell cuối cùng có dữ liệu trong cột
                Range lastCell = worksheet.Cells[worksheet.Rows.Count, columnIndex].End(XlDirection.xlUp);
                int lastRow = lastCell.Row;

                // Tạo regex pattern để match prefix + số
                // Ví dụ: エビデンス_(\d+) sẽ match エビデンス_01, エビデンス_100, etc.
                string regexPattern = $@"^{Regex.Escape(prefix)}(\d+)$";
                var regex = new Regex(regexPattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);

                if (generalConfig.EnableDebugLog)
                {
                    Logger.Debug($"Scanning column {columnIndex} from row 1 to {lastRow} for pattern '{regexPattern}'");
                }

                for (int row = 1; row <= lastRow; row++)
                {
                    try
                    {
                        var cellValue = worksheet.Cells[row, columnIndex].Value2;
                        if (cellValue != null)
                        {
                            string cellText = cellValue.ToString().Trim();

                            // Sử dụng regex để kiểm tra pattern và extract số
                            var match = regex.Match(cellText);
                            if (match.Success && match.Groups.Count > 1)
                            {
                                string numberPart = match.Groups[1].Value;
                                if (int.TryParse(numberPart, out int number))
                                {
                                    if (number > maxNumber)
                                    {
                                        maxNumber = number;
                                        if (generalConfig.EnableDebugLog)
                                        {
                                            Logger.Debug($"Found higher number: {number} in cell [{row}, {columnIndex}] = '{cellText}'");
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (generalConfig.EnableDebugLog)
                        {
                            Logger.Error($"Error reading cell [{row}, {columnIndex}]: {ex.Message}");
                        }
                    }
                }

                // Tạo tên mới với số tiếp theo
                int nextNumber = maxNumber + 1;
                string newName = "";

                // Sử dụng regex để parse format từ cấu hình (D2, D3, etc.)
                var formatRegex = new Regex(@"^D(\d+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
                var formatMatch = formatRegex.Match(numberFormat);

                if (formatMatch.Success && int.TryParse(formatMatch.Groups[1].Value, out int digits))
                {
                    // Tính số chữ số cần thiết cho nextNumber
                    int requiredDigits = nextNumber.ToString().Length;

                    // Sử dụng số chữ số lớn hơn giữa cấu hình và số cần thiết
                    int finalDigits = Math.Max(digits, requiredDigits);

                    // Tạo format string mới
                    string finalFormat = $"D{finalDigits}";
                    newName = $"{prefix}{nextNumber.ToString(finalFormat)}";

                    if (generalConfig.EnableDebugLog && finalDigits > digits)
                    {
                        Logger.Debug($"Auto-adjusted format from D{digits} to D{finalDigits} for number {nextNumber}");
                    }
                }
                else
                {
                    // Fallback về format tự động dựa trên số cần thiết
                    int requiredDigits = Math.Max(2, nextNumber.ToString().Length); // Tối thiểu 2 chữ số
                    newName = $"{prefix}{nextNumber.ToString($"D{requiredDigits}")}";

                    if (generalConfig.EnableDebugLog)
                    {
                        Logger.Debug($"Invalid format '{numberFormat}', using auto format D{requiredDigits}");
                    }
                }
                if (generalConfig.EnableDebugLog)
                {
                    Logger.Debug($"Generated auto sheet name: {newName} (max found: {maxNumber}, config: {sheetConfig.Name})");
                }

                return newName;
            }
            catch (Exception ex)
            {
                Logger.Error($"Error in GenerateAutoSheetName: {ex.Message}", ex);
                return "";
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
                        Logger.Error($"Exception in delayed SetActiveSheet: {ex.Message}");
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
                Logger.Error($"Error in SetActiveSheet: {ex.Message}");
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
                Logger.Error($"Error in Application_SheetActivate: {ex.Message}");
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
        public void TogglePinSheet(String workbookName, String sheetName)
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
        public bool IsSheetPinned(String workbookName, String sheetName)
        {
            return PinnedSheets.ContainsKey(workbookName) && PinnedSheets[workbookName].Contains(sheetName);
        }
        #endregion

        /// <summary>
        /// Load và áp dụng các template từ các đường dẫn được chỉ định, hỗ trợ các phiên bản khác nhau của Office
        /// </summary>
        public void LoadTemplate(Workbook wb)
        {
            try
            {
                // Lấy phiên bản Office hiện tại
                var app = Globals.ThisAddIn.Application;
                if (app == null)
                {
                    Logger.Error("Không thể truy cập ứng dụng Office.");
                    return;
                }

                string officeVersion = app.Version;
                string officeBasePath = "";

                // Xác định đường dẫn cơ sở dựa trên phiên bản Office
                switch (officeVersion)
                {
                    case "15.0": // Office 2013
                        officeBasePath = "C:\\Program Files\\Microsoft Office\\Office15\\Document Themes 15";
                        break;
                    case "16.0": // Office 2016, Office 365
                        officeBasePath = "C:\\Program Files\\Microsoft Office\\root\\Document Themes 16";
                        break;
                    default:
                        Logger.Error($"Phiên bản Office không được hỗ trợ: {officeVersion}");
                        return;
                }

                if (!System.IO.Directory.Exists(officeBasePath))
                {
                    Logger.Error("Không tìm thấy thư mục Document Themes. Vui lòng kiểm tra cài đặt Office.");
                    return;
                }

                string themeColorsPath = System.IO.Path.Combine(officeBasePath, "Theme Colors", "Office 2007 - 2010.xml");
                string themeFontsPath = System.IO.Path.Combine(officeBasePath, "Theme Fonts", "Office 2007 - 2010.xml");

                // Kiểm tra sự tồn tại của các tệp
                if (!System.IO.File.Exists(themeColorsPath))
                {
                    Logger.Error($"Không tìm thấy tệp Theme Colors tại: {themeColorsPath}");
                    return;
                }

                if (!System.IO.File.Exists(themeFontsPath))
                {
                    Logger.Error($"Không tìm thấy tệp Theme Fonts tại: {themeFontsPath}");
                    return;
                }

                Logger.Info("Bắt đầu áp dụng các template...");

                // Áp dụng Theme Colors
                wb.Theme.ThemeColorScheme.Load(themeColorsPath);
                Logger.Info($"Đã áp dụng Theme Colors từ: {themeColorsPath}");

                // Áp dụng Theme Fonts
                wb.Theme.ThemeFontScheme.Load(themeFontsPath);
                Logger.Info($"Đã áp dụng Theme Fonts từ: {themeFontsPath}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Có lỗi xảy ra khi áp dụng template: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Các ký tự đặc biệt cần thay thế trong tên named range
        /// </summary>
        private static readonly char[] INVALID_NAME_CHARACTERS = { '', '.', '~', ' ', '-', '+', '=', '*', '/', '\\', '[', ']', '(', ')', '{', '}', '<', '>', '!', '@', '#', '$', '%', '^', '&', '|', ':', ';', '"', '\'', ',', '?' };

        /// <summary>
        /// Tạo named range cho ô hiện tại với tên dựa trên giá trị của ô
        /// </summary>
        /// <returns>Tên named range đã tạo nếu thành công, null nếu thất bại</returns>
        private string CreateNamedRange()
        {
            try
            {
                // Validate và lấy các object cần thiết
                if (!TryGetActiveObjects(out var activeWorkbook, out var activeSheet, out var activeCell))
                {
                    return null;
                }

                // Tạo tên cho named range
                string rangeName = GenerateValidRangeName(activeCell);
                if (string.IsNullOrEmpty(rangeName))
                {
                    Logger.Warning("Không thể tạo tên named range hợp lệ từ giá trị ô");
                    return null;
                }

                // Đảm bảo tên unique
                rangeName = EnsureUniqueRangeName(activeWorkbook, rangeName);
                if (string.IsNullOrEmpty(rangeName))
                {
                    Logger.Warning("Không thể tạo tên named range unique");
                    return null;
                }

                // Tạo named range
                if (CreateNamedRangeInternal(activeWorkbook, activeSheet, activeCell, rangeName))
                {
                    return rangeName; // Trả về tên named range đã tạo
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Có lỗi xảy ra khi tạo named range: {ex.Message}", ex);
                MessageBox.Show($"Có lỗi xảy ra khi tạo named range: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        /// <summary>
        /// Lấy và validate các object Excel hiện tại
        /// </summary>
        private bool TryGetActiveObjects(out Workbook workbook, out Worksheet worksheet, out Range cell)
        {
            workbook = null;
            worksheet = null;
            cell = null;

            try
            {
                workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                worksheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;
                cell = Globals.ThisAddIn.Application.ActiveCell as Range;

                if (workbook == null)
                {
                    Logger.Warning("Không có workbook nào đang mở");
                    return false;
                }

                if (worksheet == null)
                {
                    Logger.Warning("Không có worksheet nào đang active");
                    return false;
                }

                if (cell == null)
                {
                    Logger.Warning("Không có cell nào đang được chọn");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"Lỗi khi lấy active objects: {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// Tạo tên hợp lệ cho named range từ giá trị của cell
        /// </summary>
        private string GenerateValidRangeName(Range cell)
        {
            try
            {
                string cellValue = cell.Value2?.ToString()?.Trim() ?? "";

                if (string.IsNullOrEmpty(cellValue))
                {
                    Logger.Debug("Giá trị ô trống, không thể tạo tên named range");
                    return null;
                }

                // Thay thế các ký tự không hợp lệ
                string validName = cellValue;
                foreach (char invalidChar in INVALID_NAME_CHARACTERS)
                {
                    validName = validName.Replace(invalidChar.ToString(), "_");
                }

                // Đảm bảo tên bắt đầu bằng chữ cái hoặc underscore
                if (!char.IsLetter(validName[0]) && validName[0] != '_')
                {
                    validName = "_" + validName;
                }

                // Giới hạn độ dài tên (Excel cho phép tối đa 255 ký tự)
                if (validName.Length > 255)
                {
                    validName = validName.Substring(0, 255);
                }

                // Loại bỏ các underscore liên tiếp
                while (validName.Contains("__"))
                {
                    validName = validName.Replace("__", "_");
                }

                // Loại bỏ underscore ở cuối
                validName = validName.TrimEnd('_');

                // Đảm bảo tên không rỗng sau khi clean up
                if (string.IsNullOrEmpty(validName) || validName == "_")
                {
                    Logger.Debug("Tên named range rỗng sau khi clean up, sử dụng tên mặc định");
                    validName = $"Cell_{cell.Row}_{cell.Column}";
                }

                // Validate final name
                if (!IsValidNamedRangeName(validName))
                {
                    Logger.Warning($"Tên generated '{validName}' không hợp lệ, sử dụng fallback");
                    validName = $"Cell_{cell.Row}_{cell.Column}";
                }

                return validName;
            }
            catch (Exception ex)
            {
                Logger.Error($"Lỗi khi tạo tên named range: {ex.Message}", ex);
                return null;
            }
        }

        /// <summary>
        /// Đảm bảo tên named range là duy nhất trong workbook
        /// </summary>
        private string EnsureUniqueRangeName(Workbook workbook, string baseName)
        {
            try
            {
                if (string.IsNullOrEmpty(baseName))
                {
                    Logger.Error("Base name is null or empty");
                    return null;
                }

                string uniqueName = baseName;
                int counter = 1;

                // Kiểm tra xem tên đã tồn tại chưa
                while (NameExistsInWorkbook(workbook, uniqueName))
                {
                    uniqueName = $"{baseName}_{counter}";
                    counter++;

                    // Validate tên mới
                    if (!IsValidNamedRangeName(uniqueName))
                    {
                        Logger.Warning($"Generated unique name '{uniqueName}' không hợp lệ, thử lại");
                        continue;
                    }

                    // Tránh vòng lặp vô hạn
                    if (counter > 1000)
                    {
                        Logger.Warning($"Không thể tạo tên unique sau 1000 lần thử cho base name: {baseName}");
                        return null;
                    }
                }

                // Final validation
                if (!IsValidNamedRangeName(uniqueName))
                {
                    Logger.Error($"Final unique name '{uniqueName}' không hợp lệ");
                    return null;
                }

                return uniqueName;
            }
            catch (Exception ex)
            {
                Logger.Error($"Lỗi khi kiểm tra tên unique: {ex.Message}", ex);
                return null;
            }
        }

        /// <summary>
        /// Kiểm tra xem tên đã tồn tại trong workbook chưa
        /// </summary>
        private bool NameExistsInWorkbook(Workbook workbook, string name)
        {
            try
            {
                foreach (Name existingName in workbook.Names)
                {
                    if (string.Equals(existingName.Name, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error($"Lỗi khi kiểm tra tên tồn tại: {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// Tạo named range với xử lý lỗi chi tiết
        /// </summary>
        private bool CreateNamedRangeInternal(Workbook workbook, Worksheet worksheet, Range cell, string rangeName)
        {
            try
            {
                // Validate inputs
                if (workbook == null || worksheet == null || cell == null || string.IsNullOrEmpty(rangeName))
                {
                    Logger.Error("Invalid parameters for CreateNamedRangeInternal");
                    return false;
                }

                // Validate tên named range theo quy tắc Excel
                if (!IsValidNamedRangeName(rangeName))
                {
                    Logger.Error($"Tên named range không hợp lệ: {rangeName}");
                    return false;
                }

                // Kiểm tra lại xem tên đã tồn tại chưa (double check)
                if (NameExistsInWorkbook(workbook, rangeName))
                {
                    Logger.Warning($"Named range '{rangeName}' đã tồn tại, không thể tạo lại");
                    return false;
                }

                // Tạo named range reference
                string cellAddress = $"'{worksheet.Name}'!{cell.Address[false, false, XlReferenceStyle.xlA1]}";
                Logger.Debug($"Creating named range '{rangeName}' with address: {cellAddress}");

                // Tạo named range với workbook scope
                workbook.Names.Add(rangeName, $"={cellAddress}");

                Logger.Info($"Đã tạo named range '{rangeName}' cho ô {cell.Address[false, false]} trong sheet '{worksheet.Name}'");
                return true;
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                Logger.Error($"Lỗi COM khi tạo named range '{rangeName}': 0x{comEx.ErrorCode:X8} - {comEx.Message}", comEx);

                // Xử lý các lỗi cụ thể
                switch (comEx.ErrorCode)
                {
                    case -2146827284: // 0x800A03EC - General Excel error
                        Logger.Error("Lỗi Excel chung - có thể do tên không hợp lệ hoặc workbook corrupted");
                        break;
                    case -2147352567: // 0x80020009 - Invalid name
                        Logger.Error("Tên named range không hợp lệ theo quy tắc Excel");
                        break;
                    case -2146827864: // 0x800A01A8 - Name already exists
                        Logger.Error("Named range đã tồn tại");
                        break;
                    default:
                        Logger.Error($"Lỗi COM không xác định: 0x{comEx.ErrorCode:X8}");
                        break;
                }

                return false;
            }
            catch (Exception ex)
            {
                Logger.Error($"Lỗi không xác định khi tạo named range '{rangeName}': {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// Kiểm tra xem tên named range có hợp lệ theo quy tắc Excel không
        /// </summary>
        private bool IsValidNamedRangeName(string name)
        {
            try
            {
                if (string.IsNullOrEmpty(name))
                    return false;

                // Tên không được bắt đầu bằng số
                if (char.IsDigit(name[0]))
                    return false;

                // Tên không được chứa khoảng trắng ở đầu hoặc cuối
                if (name != name.Trim())
                    return false;

                // Tên không được chứa các ký tự đặc biệt (ngoài underscore)
                foreach (char c in name)
                {
                    if (!char.IsLetterOrDigit(c) && c != '_')
                        return false;
                }

                // Tên không được là địa chỉ cell (như A1, B2, etc.)
                if (System.Text.RegularExpressions.Regex.IsMatch(name, @"^[A-Z]+\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    return false;

                // Tên không được vượt quá 255 ký tự
                if (name.Length > 255)
                    return false;

                // Tên không được là các từ khóa Excel reserved
                string[] reservedWords = { "PRINT_AREA", "PRINT_TITLES", "FILTER_DATABASE", "EXTRACT", "CONSOLIDATE_AREA", "DATABASE", "CRITERIA" };
                if (Array.Exists(reservedWords, word => string.Equals(word, name, StringComparison.OrdinalIgnoreCase)))
                    return false;

                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"Lỗi khi validate tên named range: {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// Tạo hoặc lấy named range cho cell (tái sử dụng nếu đã tồn tại)
        /// </summary>
        private string GetOrCreateNamedRangeForCell(Range cell, Worksheet sourceSheet)
        {
            try
            {
                // Lưu trữ active cell hiện tại
                var originalActiveCell = Globals.ThisAddIn.Application.ActiveCell as Range;

                // Activate source sheet và cell
                sourceSheet.Activate();
                cell.Select();

                // Tạo named range và lấy tên
                string namedRangeName = CreateNamedRange();

                // Khôi phục active cell gốc nếu có
                if (originalActiveCell != null)
                {
                    try
                    {
                        originalActiveCell.Worksheet.Activate();
                        originalActiveCell.Select();
                    }
                    catch { /* Ignore errors when restoring */ }
                }

                return namedRangeName;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Không thể tạo/lấy named range cho cell: {ex.Message}");
                return null;
            }
        }
    }
}