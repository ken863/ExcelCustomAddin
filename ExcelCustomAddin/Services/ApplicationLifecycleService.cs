using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Threading;

namespace ExcelCustomAddin
{
  /// <summary>
  /// ApplicationLifecycleService - Quản lý vòng đời ứng dụng Excel Add-in
  ///
  /// Chức năng chính:
  /// - Khởi động và tắt add-in an toàn
  /// - Phát hiện phiên bản Excel và kiến trúc hệ thống
  /// - Quản lý theme và template cho workbook
  /// - Xử lý sự kiện Excel (workbook open/close/activate, sheet changes)
  /// - Quản lý Action Panel cho từng workbook
  /// - Đăng ký/hủy đăng ký events
  ///
  /// Kiến trúc hỗ trợ:
  /// - Excel 2013+ (Office 15.0, 16.0)
  /// - Windows x64 và x86
  /// - Theme Office 2007-2010
  ///
  /// Tác giả: lam.pt
  /// Ngày tạo: 2025
  /// </summary>
  public class ApplicationLifecycleService
  {
    #region Fields

    private readonly ThisAddIn _addIn;

    // Thông tin phiên bản Excel và kiến trúc hệ thống
    private readonly string _excelVersion;
    private readonly string _architecture;
    private readonly bool _isExcel64Bit;

    #endregion

    #region Constructor

    /// <summary>
    /// Khởi tạo ApplicationLifecycleService
    /// Phát hiện và lưu trữ thông tin về phiên bản Excel và kiến trúc hệ thống
    ///
    /// </summary>
    /// <param name="addIn">Instance của ThisAddIn chính</param>
    public ApplicationLifecycleService(ThisAddIn addIn)
    {
      _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));

      // Phát hiện phiên bản Excel và kiến trúc hệ thống
      _excelVersion = GetExcelVersion();
      _architecture = GetSystemArchitecture();
      _isExcel64Bit = IsExcel64Bit();

      Logger.Info($"ApplicationLifecycleService initialized - Excel: {_excelVersion}, Architecture: {_architecture}, Excel x64: {_isExcel64Bit}");
    }

    #endregion

    #region Public Interface

    /// <summary>
    /// Khởi động add-in và thiết lập môi trường làm việc
    ///
    /// Quy trình khởi động:
    /// 1. Load cấu hình từ XML
    /// 2. Thiết lập logging
    /// 3. Tạo Dispatcher cho UI thread
    /// 4. Đăng ký application events
    /// 5. Khởi tạo Action Panel cho workbook hiện tại
    ///
    /// </summary>
    public void Startup()
    {
      Logger.Info("=== Excel Custom Add-in Startup Sequence ===");

      try
      {
        // Bước 1: Load cấu hình từ SheetConfigManager
        Logger.Info("Loading sheet configuration...");
        SheetConfigManager.LoadConfiguration();
        Logger.Info("✓ Sheet configuration loaded successfully");

        // Bước 2: Thiết lập logging configuration
        var loggingConfig = SheetConfigManager.GetLoggingConfig();
        Logger.Info($"Logger configured - Directory: {(string.IsNullOrEmpty(loggingConfig.LogDirectory) ? "Default C:\\ExcelCustomAddin" : loggingConfig.LogDirectory)}, File: {loggingConfig.LogFileName}, Debug: {loggingConfig.EnableDebugOutput}");
        Logger.Info($"Log file path: {Logger.GetLogFilePath()}");

        // Bước 3: Cấu hình debug logging
        var generalConfig = SheetConfigManager.GetGeneralConfig();
        if (generalConfig != null)
        {
          Logger.SetDebugEnabled(generalConfig.EnableDebugLog);
          Logger.Debug($"Debug logging {(generalConfig.EnableDebugLog ? "enabled" : "disabled")}");
        }

        // Bước 4: Thiết lập UI thread dispatcher
        _addIn._dispatcher = Dispatcher.CurrentDispatcher;
        Logger.Debug("UI dispatcher initialized");

        // Bước 5: Đăng ký application events
        RegisterApplicationEvents();
        Logger.Info("✓ Application events registered");

        // Bước 6: Khởi tạo Action Panel với delay để tránh xung đột
        if (_addIn.Application.ActiveWorkbook != null)
        {
          Logger.Debug("Initializing Action Panel for active workbook...");
          var timer = new System.Windows.Forms.Timer { Interval = 500 };
          timer.Tick += (s, args) =>
          {
            timer.Stop();
            timer.Dispose();
            CreateActionsPane(_addIn.Application.ActiveWorkbook);
          };
          timer.Start();
        }

        Logger.Info("=== Startup sequence completed successfully ===");
      }
      catch (Exception ex)
      {
        Logger.Error("Critical error during startup", ex);
        throw; // Re-throw để caller xử lý
      }
    }

    /// <summary>
    /// Tắt add-in và dọn dẹp tài nguyên
    ///
    /// Quy trình tắt:
    /// 1. Hủy đăng ký application events
    /// 2. Hủy đăng ký Action Panel events
    /// 3. Log completion
    ///
    /// </summary>
    public void Shutdown()
    {
      Logger.Info("=== Excel Custom Add-in Shutdown Sequence ===");

      try
      {
        // Hủy đăng ký application events
        UnregisterApplicationEvents();
        Logger.Info("✓ Application events unregistered");

        // Hủy đăng ký Action Panel events
        UnregisterActionPanelEvents();
        Logger.Info("✓ Action Panel events unregistered");

        Logger.Info("=== Shutdown sequence completed successfully ===");
      }
      catch (Exception ex)
      {
        Logger.Error("Error during shutdown", ex);
        // Không re-throw trong shutdown để tránh crash
      }
    }

    /// <summary>
    /// Load và áp dụng các template từ các đường dẫn được chỉ định, hỗ trợ các phiên bản khác nhau của Office
    /// </summary>
    /// <param name="wb">Workbook cần áp dụng theme</param>
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
        bool foundPath = false;

        // Lấy các đường dẫn Program Files (cả x64 và x86)
        var programFilesPaths = new[]
        {
          Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),      // Program Files (x64)
          Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)    // Program Files (x86)
        };

        // Xác định đường dẫn cơ sở dựa trên phiên bản Office và kiểm tra cả x64/x86
        foreach (var programFiles in programFilesPaths)
        {
          if (string.IsNullOrEmpty(programFiles)) continue;

          switch (officeVersion)
          {
            case "15.0": // Office 2013
              officeBasePath = System.IO.Path.Combine(programFiles, "Microsoft Office", "Office15", "Document Themes 15");
              break;
            case "16.0": // Office 2016, Office 365
              officeBasePath = System.IO.Path.Combine(programFiles, "Microsoft Office", "root", "Document Themes 16");
              break;
            default:
              Logger.Error($"Phiên bản Office không được hỗ trợ: {officeVersion}");
              return;
          }

          if (System.IO.Directory.Exists(officeBasePath))
          {
            foundPath = true;
            Logger.Debug($"Tìm thấy thư mục Document Themes tại: {officeBasePath}");
            break;
          }
        }

        if (!foundPath)
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

    #endregion

    #region Excel Version Detection

    /// <summary>
    /// Lấy phiên bản Excel hiện tại
    /// Sử dụng Application.Version property của Excel
    ///
    /// </summary>
    /// <returns>Chuỗi phiên bản Excel (VD: "Excel 16.0")</returns>
    private string GetExcelVersion()
    {
      try
      {
        var version = Globals.ThisAddIn.Application.Version;
        return $"Excel {version}";
      }
      catch
      {
        return "Unknown";
      }
    }

    /// <summary>
    /// Lấy kiến trúc của hệ thống (.NET runtime)
    /// Xác định xem process hiện tại là 64-bit hay 32-bit
    ///
    /// </summary>
    /// <returns>"x64" hoặc "x86"</returns>
    private string GetSystemArchitecture() => Environment.Is64BitProcess ? "x64" : "x86";

    /// <summary>
    /// Kiểm tra Excel có phải phiên bản 64-bit không
    /// Sử dụng Marshal.SizeOf để kiểm tra pointer size
    ///
    /// </summary>
    /// <returns>true nếu Excel 64-bit, false nếu 32-bit</returns>
    private bool IsExcel64Bit()
    {
      try
      {
        // Kiểm tra size của IntPtr để xác định bitness
        return Marshal.SizeOf(IntPtr.Zero) == 8;
      }
      catch
      {
        // Fallback: kiểm tra process
        return Environment.Is64BitProcess;
      }
    }

    #endregion

    #region Event Registration

    /// <summary>
    /// Đăng ký tất cả application events cần thiết
    /// Bao gồm workbook và worksheet events
    ///
    /// Events được đăng ký:
    /// - Workbook: NewWorkbook, WorkbookOpen, WorkbookActivate, WorkbookBeforeClose, WorkbookAfterSave
    /// - Worksheet: SheetActivate
    ///
    /// </summary>
    private void RegisterApplicationEvents()
    {
      var app = Globals.ThisAddIn.Application;

      // Workbook events
      ((AppEvents_Event)app).NewWorkbook += Application_NewWorkbook;
      app.WorkbookOpen += Application_WorkbookOpen;
      app.WorkbookActivate += Application_WorkbookActive;
      app.WorkbookBeforeClose += Application_WorkbookBeforeClose;
      app.WorkbookAfterSave += Application_WorkbookAfterSave;

      // Worksheet events
      app.SheetActivate += Application_SheetActivate;

      Logger.Debug("Application events registered successfully");
    }

    /// <summary>
    /// Hủy đăng ký tất cả application events
    /// Đảm bảo cleanup proper khi shutdown
    ///
    /// </summary>
    private void UnregisterApplicationEvents()
    {
      var app = Globals.ThisAddIn.Application;
      if (app != null)
      {
        // Workbook events
        ((AppEvents_Event)app).NewWorkbook -= Application_NewWorkbook;
        app.WorkbookOpen -= Application_WorkbookOpen;
        app.WorkbookActivate -= Application_WorkbookActive;
        app.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
        app.WorkbookAfterSave -= Application_WorkbookAfterSave;

        // Worksheet events
        app.SheetActivate -= Application_SheetActivate;

        Logger.Debug("Application events unregistered successfully");
      }
    }

    /// <summary>
    /// Đăng ký events cho ActionPanel control
    /// Liên kết các button events với handlers
    ///
    /// </summary>
    /// <param name="actionPanel">ActionPanel control cần đăng ký events</param>
    private void RegisterActionPanelEvents(ActionPanelControl actionPanel)
    {
      actionPanel.CreateEvidenceEvent += _addIn.CreateEvidence;
      actionPanel.FormatImagesEvent += _addIn.FormatImages;
      actionPanel.FormatDocumentEvent += _addIn.FormatDocument;
      actionPanel.ChangeSheetNameEvent += _addIn.ChangeSheetName;
      actionPanel.InsertMultipleImagesEvent += _addIn.InsertMultipleImages;
      actionPanel.PinSheetEvent += _addIn.PinSheet;
      actionPanel.listofSheet.SelectedIndexChanged += _addIn._sheetService.ListOfSheet_SelectionChanged;
    }

    /// <summary>
    /// Hủy đăng ký events cho ActionPanel control
    /// Cleanup trước khi tạo ActionPanel mới
    ///
    /// </summary>
    /// <param name="actionPanel">ActionPanel control cần hủy đăng ký</param>
    private void UnregisterActionPanelEvents(ActionPanelControl actionPanel)
    {
      actionPanel.CreateEvidenceEvent -= _addIn.CreateEvidence;
      actionPanel.FormatImagesEvent -= _addIn.FormatImages;
      actionPanel.FormatDocumentEvent -= _addIn.FormatDocument;
      actionPanel.ChangeSheetNameEvent -= _addIn.ChangeSheetName;
      actionPanel.InsertMultipleImagesEvent -= _addIn.InsertMultipleImages;
      actionPanel.PinSheetEvent -= _addIn.PinSheet;
      actionPanel.listofSheet.SelectedIndexChanged -= _addIn._sheetService.ListOfSheet_SelectionChanged;
    }

    /// <summary>
    /// Hủy đăng ký ActionPanel events (overload)
    /// Sử dụng khi không có ActionPanel instance cụ thể
    ///
    /// </summary>
    private void UnregisterActionPanelEvents()
    {
      if (_addIn._actionPanel != null)
        UnregisterActionPanelEvents(_addIn._actionPanel);
    }

    #endregion

    #region Application Event Handlers

    /// <summary>
    /// Xử lý sự kiện tạo workbook mới
    /// Trigger: File > New hoặc Ctrl+N
    ///
    /// </summary>
    /// <param name="wb">Workbook mới được tạo</param>
    private void Application_NewWorkbook(Workbook wb) => HandleWorkbookOpened(wb, "NewWorkbook");

    /// <summary>
    /// Xử lý sự kiện mở workbook
    /// Trigger: File > Open hoặc double-click file Excel
    ///
    /// </summary>
    /// <param name="wb">Workbook được mở</param>
    private void Application_WorkbookOpen(Workbook wb) => HandleWorkbookOpened(wb, "WorkbookOpen");

    /// <summary>
    /// Handler chung cho các sự kiện workbook opened
    /// Thực hiện các tác vụ chung: load template và tạo Action Panel
    ///
    /// </summary>
    /// <param name="wb">Workbook được mở</param>
    /// <param name="eventName">Tên event để logging</param>
    private void HandleWorkbookOpened(Workbook wb, string eventName)
    {
      Logger.Debug($"Application_{eventName} called for: {wb?.Name}");

      if (wb != null)
      {
        LoadTemplate(wb);
        CreateActionsPane(wb);
      }
    }

    /// <summary>
    /// Xử lý sự kiện activate workbook
    /// Trigger: Click tab workbook hoặc Alt+Tab
    ///
    /// </summary>
    /// <param name="wb">Workbook được activate</param>
    private void Application_WorkbookActive(Workbook wb)
    {
      if (wb != null && ThisAddIn.CreatedActionPanes.Contains(wb.Name))
      {
        UpdateActionPanelForWorkbook(wb, wb.Name);
        Logger.Debug($"Updated Action Panel for activated workbook: {wb.Name}");
      }
    }

    /// <summary>
    /// Xử lý sự kiện trước khi đóng workbook
    /// Trigger: File > Close hoặc click X trên workbook
    ///
    /// </summary>
    /// <param name="wb">Workbook sắp được đóng</param>
    /// <param name="cancel">Có thể set true để hủy việc đóng</param>
    private void Application_WorkbookBeforeClose(Workbook wb, ref bool cancel)
    {
      if (wb != null)
      {
        CleanupWorkbookResources(wb.Name);
        Logger.Debug($"Cleaned up resources for workbook: {wb.Name}");
      }
    }

    /// <summary>
    /// Xử lý sự kiện sau khi lưu workbook
    /// Trigger: File > Save hoặc Ctrl+S
    ///
    /// </summary>
    /// <param name="wb">Workbook đã được lưu</param>
    /// <param name="success">true nếu lưu thành công</param>
    private void Application_WorkbookAfterSave(Workbook wb, bool success)
    {
      if (wb != null && success && _addIn._actionPanel != null)
      {
        _addIn._actionPanel.RefreshFilePathDisplay();
        Logger.Debug($"Refreshed file path display after save for: {wb.Name}");
      }
    }

    /// <summary>
    /// Xử lý sự kiện activate worksheet
    /// Trigger: Click tab sheet hoặc chuyển sheet bằng code
    ///
    /// </summary>
    /// <param name="sh">Worksheet được activate (có thể là Chart, etc.)</param>
    private void Application_SheetActivate(object sh)
    {
      try
      {
        var worksheet = sh as Worksheet;
        if (worksheet == null) return;

        Logger.Debug($"Sheet activated: {worksheet.Name}");

        // Đánh dấu sheet đang được activate để tránh recursive updates
        _addIn.IsSheetActivating = true;

        // Cập nhật danh sách sheet trong Action Panel
        _addIn._actionPanel?.BindSheetList(_addIn._sheetService.GetListOfSheet(), worksheet.Name);

        // Reset flag sau delay ngắn
        var timer = new System.Windows.Forms.Timer { Interval = 100 };
        timer.Tick += (s, args) =>
        {
          timer.Stop();
          timer.Dispose();
          _addIn.IsSheetActivating = false;
        };
        timer.Start();
      }
      catch (Exception ex)
      {
        Logger.Error($"Error in Application_SheetActivate: {ex.Message}", ex);
      }
    }

    #endregion

    #region Action Panel Management

    /// <summary>
    /// Tạo Action Panel cho workbook
    /// Quản lý việc tạo mới hoặc cập nhật Action Panel hiện có
    ///
    /// Logic:
    /// 1. Kiểm tra xem Action Panel đã tồn tại cho workbook chưa
    /// 2. Nếu có: chỉ cập nhật
    /// 3. Nếu chưa: tạo mới với đầy đủ event registration
    ///
    /// </summary>
    /// <param name="wb">Workbook cần tạo Action Panel</param>
    private void CreateActionsPane(Workbook wb)
    {
      if (wb == null) return;

      string workbookKey = wb.Name;

      lock (ThisAddIn._lockObject)
      {
        Logger.Debug($"CreateActionsPane called for: {workbookKey}");

        if (ThisAddIn.CreatedActionPanes.Contains(workbookKey))
        {
          // Action Panel đã tồn tại - chỉ cập nhật
          Logger.Debug($"Action panel already exists for: {workbookKey}, updating only");
          UpdateActionPanelForWorkbook(wb, workbookKey);
          return;
        }

        // Tạo Action Panel mới
        Logger.Debug($"Creating new action panel for: {workbookKey}");
        CreateNewActionPanel(wb, workbookKey);
      }
    }

    /// <summary>
    /// Cập nhật Action Panel cho workbook hiện có
    /// Chỉ cập nhật danh sách sheet và focus, không tạo lại UI
    ///
    /// </summary>
    /// <param name="wb">Workbook cần cập nhật</param>
    /// <param name="workbookKey">Key của workbook</param>
    private void UpdateActionPanelForWorkbook(Workbook wb, string workbookKey)
    {
      var currentTaskPane = TaskPaneManager.GetTaskPane(workbookKey, "WORKSHEET TOOLS", null);
      if (currentTaskPane != null)
      {
        _addIn.myCustomTaskPane = currentTaskPane;
        _addIn._actionPanel = (ActionPanelControl)_addIn.myCustomTaskPane.Control;

        var currentSheetName = wb.ActiveSheet?.Name;
        _addIn._actionPanel.BindSheetList(_addIn._sheetService.GetListOfSheet(), currentSheetName);
      }
    }

    /// <summary>
    /// Tạo Action Panel mới hoàn toàn cho workbook
    /// Bao gồm tạo UI, đăng ký events và cấu hình
    ///
    /// </summary>
    /// <param name="wb">Workbook cần tạo Action Panel</param>
    /// <param name="workbookKey">Key của workbook</param>
    private void CreateNewActionPanel(Workbook wb, string workbookKey)
    {
      _addIn.myCustomTaskPane = TaskPaneManager.GetTaskPane(wb.Name, "WORKSHEET TOOLS", () => new ActionPanelControl());
      _addIn._actionPanel = (ActionPanelControl)_addIn.myCustomTaskPane?.Control;

      if (_addIn._actionPanel != null)
      {
        // Đăng ký events (hủy đăng ký cũ trước nếu có)
        UnregisterActionPanelEvents(_addIn._actionPanel);
        RegisterActionPanelEvents(_addIn._actionPanel);

        // Cập nhật danh sách sheet
        var currentSheetName = wb.ActiveSheet?.Name;
        _addIn._actionPanel.BindSheetList(_addIn._sheetService.GetListOfSheet(), currentSheetName);

        // Cấu hình UI
        _addIn.myCustomTaskPane.Visible = true;
        _addIn.myCustomTaskPane.Width = 300;

        // Đánh dấu workbook đã có Action Panel
        ThisAddIn.CreatedActionPanes.Add(workbookKey);
        Logger.Debug($"Action panel created and marked for: {workbookKey}");
      }
    }

    /// <summary>
    /// Dọn dẹp tài nguyên của workbook khi đóng
    /// Xóa khỏi danh sách created panes và cleanup pinned sheets
    ///
    /// </summary>
    /// <param name="workbookKey">Key của workbook cần cleanup</param>
    private void CleanupWorkbookResources(string workbookKey)
    {
      ThisAddIn.CreatedActionPanes.Remove(workbookKey);
      ThisAddIn.PinnedSheets.Remove(workbookKey);
      Logger.Debug($"Cleaned up resources for workbook: {workbookKey}");
    }

    #endregion
  }
}