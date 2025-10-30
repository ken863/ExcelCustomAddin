namespace ExcelCustomAddin
{
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using System;
    using System.Collections.Generic;
    using System.Windows;
    using System.Windows.Threading;

    /// <summary>
    /// ThisAddIn - Lớp chính của Excel Add-in
    /// Chức năng: Orchestrator cho tất cả các service, quản lý lifecycle của add-in
    /// </summary>
    public partial class ThisAddIn
    {
        #region Fields - Private và Static

        /// <summary>
        /// Action Panel Control - UI control cho task pane
        /// </summary>
        internal ActionPanelControl _actionPanel { get; set; }

        /// <summary>
        /// Custom Task Pane - Container cho action panel
        /// </summary>
        public CustomTaskPane myCustomTaskPane { get; set; }

        /// <summary>
        /// Dispatcher - Quản lý UI thread operations
        /// </summary>
        internal Dispatcher _dispatcher;

        /// <summary>
        /// Flag kiểm tra sheet đang trong quá trình kích hoạt
        /// Ngăn chặn xung đột giữa SelectedIndexChanged và SheetActivate events
        /// </summary>
        internal bool IsSheetActivating { get; set; } = false;

        /// <summary>
        /// Dictionary lưu trữ danh sách các sheet được pin theo workbook
        /// Key: Tên workbook, Value: HashSet các tên sheet được pin
        /// </summary>
        internal static Dictionary<string, HashSet<string>> PinnedSheets
            = new Dictionary<string, HashSet<string>>();

        /// <summary>
        /// HashSet lưu trữ danh sách workbook đã được tạo action panel
        /// Tránh tạo trùng lặp action panel cho cùng một workbook
        /// </summary>
        internal static HashSet<string> CreatedActionPanes
            = new HashSet<string>();

        /// <summary>
        /// Lock object đảm bảo thread safety khi tạo/cập nhật action panel
        /// </summary>
        internal static readonly object _lockObject = new object();

        #endregion

        #region Services - Dependency Injection Pattern

        /// <summary>
        /// Service quản lý lifecycle của application (startup, shutdown, workbook events)
        /// </summary>
        private ApplicationLifecycleService _lifecycleService;

        /// <summary>
        /// Service xử lý các thao tác liên quan đến hình ảnh (format, insert, scale)
        /// </summary>
        internal ImageProcessingService _imageService;

        /// <summary>
        /// Service tạo và quản lý Evidence sheets
        /// </summary>
        internal EvidenceCreationService _evidenceService;

        /// <summary>
        /// Service quản lý các thao tác liên quan đến sheet (list, select, pin, format)
        /// </summary>
        internal SheetManagementService _sheetService;

        #endregion

        #region Configuration Properties

        /// <summary>
        /// Tên cột được sử dụng làm page break / cột ngoài cùng bên phải khi in
        /// Được load từ SheetConfig.xml
        /// </summary>
        private static string PAGE_BREAK_COLUMN_NAME => SheetConfigManager.GetGeneralConfig().PageBreakColumnName;

        /// <summary>
        /// Font mặc định cho Evidence sheets và các cell chứa hyperlink
        /// Được load từ SheetConfig.xml
        /// </summary>
        private static string EVIDENCE_FONT_NAME => SheetConfigManager.GetGeneralConfig().EvidenceFontName;

        /// <summary>
        /// Chỉ số dòng cuối cùng trong vùng in của Evidence sheets
        /// Được load từ SheetConfig.xml
        /// </summary>
        private static int PRINT_AREA_LAST_ROW_IDX => SheetConfigManager.GetGeneralConfig().PrintAreaLastRowIdx;

        #endregion

        #region Lifecycle Methods

        /// <summary>
        /// InternalStartup - Đăng ký event handlers cho Startup và Shutdown
        /// Được gọi tự động bởi VSTO runtime
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        /// <summary>
        /// ThisAddIn_Startup - Khởi tạo add-in khi Excel khởi động
        /// </summary>
        /// <param name="sender">Object gọi event</param>
        /// <param name="e">Event arguments</param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Khởi tạo tất cả các service
            InitializeServices();

            // Gọi Startup của lifecycle service để đăng ký events và load configuration
            _lifecycleService.Startup();
        }

        /// <summary>
        /// ThisAddIn_Shutdown - Dọn dẹp tài nguyên khi Excel đóng
        /// </summary>
        /// <param name="sender">Object gọi event</param>
        /// <param name="e">Event arguments</param>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Delegate cleanup cho lifecycle service
            _lifecycleService?.Shutdown();
        }

        /// <summary>
        /// InitializeServices - Khởi tạo tất cả các service instances
        /// Sử dụng dependency injection pattern, truyền 'this' để services có thể truy cập ThisAddIn
        /// </summary>
        private void InitializeServices()
        {
            _lifecycleService = new ApplicationLifecycleService(this);
            _imageService = new ImageProcessingService(this);
            _evidenceService = new EvidenceCreationService(this);
            _sheetService = new SheetManagementService(this);
        }

        #endregion

        #region Event Handlers - UI Events từ ActionPanelControl

        /// <summary>
        /// Event handler cho nút Format Images
        /// Delegate thực thi cho ImageProcessingService
        /// </summary>
        /// <param name="sender">Object gọi event (ActionPanelControl)</param>
        /// <param name="e">Event arguments</param>
        internal void FormatImages(object sender, EventArgs e)
        {
            _imageService.FormatImages();
        }

        /// <summary>
        /// Event handler cho nút Insert Multiple Images
        /// Delegate thực thi cho ImageProcessingService
        /// </summary>
        /// <param name="sender">Object gọi event (ActionPanelControl)</param>
        /// <param name="e">Event arguments</param>
        internal void InsertMultipleImages(object sender, EventArgs e)
        {
            _imageService.InsertMultipleImages();
        }

        /// <summary>
        /// Event handler cho nút Format Document
        /// Delegate thực thi cho SheetManagementService
        /// </summary>
        /// <param name="sender">Object gọi event (ActionPanelControl)</param>
        /// <param name="e">Event arguments</param>
        internal void FormatDocument(object sender, EventArgs e)
        {
            _sheetService.FormatDocument();
        }

        /// <summary>
        /// Event handler cho chức năng đổi tên sheet
        /// Delegate thực thi cho SheetManagementService
        /// </summary>
        /// <param name="sender">Object gọi event (ActionPanelControl)</param>
        /// <param name="e">Event arguments</param>
        internal void ChangeSheetName(object sender, EventArgs e)
        {
            _sheetService.ChangeSheetName();
        }

        /// <summary>
        /// Event handler cho chức năng pin/unpin sheet
        /// Xử lý validation workbook trước khi delegate cho SheetManagementService
        /// </summary>
        /// <param name="sender">Object gọi event (ActionPanelControl)</param>
        /// <param name="e">Event arguments chứa tên sheet cần pin/unpin</param>
        internal void PinSheet(object sender, ActionPanelControl.PinSheetEventArgs e)
        {
            try
            {
                // Kiểm tra có workbook đang mở không
                var activeWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (activeWorkbook == null)
                {
                    Logger.Error("Không có workbook nào đang mở trong PinSheet");
                    MessageBox.Show("Không có workbook nào đang mở.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Lấy tên workbook và delegate cho service
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
        /// Event handler cho nút Create Evidence
        /// Delegate thực thi cho EvidenceCreationService
        /// </summary>
        /// <param name="sender">Object gọi event (ActionPanelControl)</param>
        /// <param name="e">Event arguments</param>
        internal void CreateEvidence(object sender, EventArgs e)
        {
            _evidenceService.CreateEvidence();
        }

        #endregion

        #region Public Methods - API cho các services và external callers

        /// <summary>
        /// Toggle pin/unpin status của một sheet trong workbook
        /// Public API wrapper cho SheetManagementService
        /// </summary>
        /// <param name="workbookName">Tên workbook chứa sheet</param>
        /// <param name="sheetName">Tên sheet cần toggle pin status</param>
        public void TogglePinSheet(String workbookName, String sheetName)
        {
            _sheetService.TogglePinSheet(workbookName, sheetName);
        }

        /// <summary>
        /// Kiểm tra xem một sheet có đang được pin hay không
        /// Public API wrapper cho SheetManagementService
        /// </summary>
        /// <param name="workbookName">Tên workbook chứa sheet</param>
        /// <param name="sheetName">Tên sheet cần kiểm tra</param>
        /// <returns>True nếu sheet đang được pin, ngược lại False</returns>
        public bool IsSheetPinned(String workbookName, String sheetName)
        {
            return _sheetService.IsSheetPinned(workbookName, sheetName);
        }

        /// <summary>
        /// Load và áp dụng template (*.dotm, *.xltm) cho workbook
        /// Hỗ trợ nhiều phiên bản Office khác nhau
        /// Public API wrapper cho ApplicationLifecycleService
        /// </summary>
        /// <param name="wb">Workbook cần áp dụng template</param>
        public void LoadTemplate(Workbook wb)
        {
            _lifecycleService.LoadTemplate(wb);
        }

        #endregion
    }
}