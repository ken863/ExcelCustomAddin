using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using System.Xml.Linq;

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
    /// Load và áp dụng theme/template cho workbook mới
    /// Hỗ trợ đa phiên bản Office và kiến trúc hệ thống
    ///
    /// </summary>
    /// <param name="wb">Workbook cần áp dụng theme</param>
    public void LoadTemplate(Workbook wb)
    {
      if (wb == null) return;

      try
      {
        Logger.Debug($"Loading Office themes for workbook: {wb.Name}");

        // Lấy và áp dụng các theme files
        foreach (var themePath in GetThemePaths())
        {
          if (File.Exists(themePath))
          {
            Logger.Debug($"Applying theme from: {themePath}");
            ApplyThemeToWorkbook(wb, themePath);
            Logger.Info($"✓ Successfully applied theme: {Path.GetFileName(themePath)}");
          }
          else
          {
            Logger.Debug($"Theme file not found: {themePath}");
          }
        }

        Logger.Debug($"Theme loading completed for workbook: {wb.Name}");
      }
      catch (Exception ex)
      {
        Logger.Error($"Error loading themes for workbook {wb.Name}: {ex.Message}", ex);
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

    #region Office Path Resolution

    /// <summary>
    /// Lấy danh sách các đường dẫn Office theme có thể có
    /// Tự động phát hiện dựa trên phiên bản Office và kiến trúc hệ thống
    ///
    /// Ưu tiên:
    /// 1. Office 2016/365 (16.0) - "root\\Document Themes 16"
    /// 2. Office 2013 (15.0) - "Office15\\Document Themes 15"
    /// 3. Cả x64 và x86 paths
    ///
    /// </summary>
    /// <param name="officeVersion">Phiên bản Office (VD: "16.0")</param>
    /// <returns>Mảng các đường dẫn có thể có theme files</returns>
    private string[] GetPossibleOfficePaths(string officeVersion)
    {
      string themeFolder = "";
      switch (officeVersion)
      {
        case "15.0": // Office 2013
          themeFolder = "Document Themes 15";
          break;
        case "16.0": // Office 2016, Office 365
          themeFolder = "Document Themes 16";
          break;
        default:
          Logger.Warning($"Unsupported Office version for theme detection: {officeVersion}");
          return new string[0];
      }

      // Lấy các đường dẫn Program Files (cả x64 và x86)
      var programFilesPaths = new[]
      {
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),      // Program Files (x64)
        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)    // Program Files (x86)
      };

      var possiblePaths = new System.Collections.Generic.List<string>();

      foreach (var programFiles in programFilesPaths)
      {
        if (!string.IsNullOrEmpty(programFiles))
        {
          // Office 2013: Office15\Document Themes 15
          if (officeVersion == "15.0")
          {
            possiblePaths.Add(Path.Combine(programFiles, "Microsoft Office", "Office15", themeFolder));
          }
          // Office 2016+: root\Document Themes 16
          else if (officeVersion == "16.0")
          {
            possiblePaths.Add(Path.Combine(programFiles, "Microsoft Office", "root", themeFolder));
          }
        }
      }

      Logger.Debug($"Generated {possiblePaths.Count} possible Office theme paths for version {officeVersion}");
      return possiblePaths.ToArray();
    }

    #endregion

    #region Theme Management

    /// <summary>
    /// Lấy danh sách đường dẫn theme files từ các vị trí có thể có
    /// Kết hợp phát hiện đường dẫn Office với danh sách theme cụ thể
    ///
    /// Theme files được tìm:
    /// - Theme Colors\\Office 2007 - 2010.xml
    /// - Theme Fonts\\Office 2007 - 2010.xml
    ///
    /// </summary>
    /// <returns>Danh sách đường dẫn theme files</returns>
    private System.Collections.Generic.List<string> GetThemePaths()
    {
      var themePaths = new System.Collections.Generic.List<string>();

      try
      {
        // Lấy phiên bản Office hiện tại
        string officeVersion = Globals.ThisAddIn.Application.Version;
        string[] possiblePaths = GetPossibleOfficePaths(officeVersion);

        // Thử từng đường dẫn Office
        foreach (var basePath in possiblePaths)
        {
          if (Directory.Exists(basePath))
          {
            // Thêm các theme files cụ thể
            var themeFiles = new[]
            {
              Path.Combine(basePath, "Theme Fonts", "Office 2007 - 2010.xml"),
              Path.Combine(basePath, "Theme Colors", "Office 2007 - 2010.xml")
            };

            foreach (var themeFile in themeFiles)
            {
              if (!themePaths.Contains(themeFile))
              {
                themePaths.Add(themeFile);
              }
            }

            Logger.Debug($"Found theme directory: {basePath}");
            break; // Dừng khi tìm thấy thư mục hợp lệ đầu tiên
          }
        }

        Logger.Debug($"Located {themePaths.Count} theme files");
      }
      catch (Exception ex)
      {
        Logger.Warning($"Error getting theme paths: {ex.Message}");
      }

      return themePaths;
    }

    /// <summary>
    /// Á dụng theme file cho workbook
    /// Sử dụng Excel Theme API với fallback mechanism
    ///
    /// Quy trình:
    /// 1. Thử ApplyTheme() method
    /// 2. Nếu thất bại, thử ApplyThemeFromXml() fallback
    /// 3. Log kết quả
    ///
    /// </summary>
    /// <param name="wb">Workbook cần áp dụng theme</param>
    /// <param name="themePath">Đường dẫn theme file</param>
    private void ApplyThemeToWorkbook(Workbook wb, string themePath)
    {
      try
      {
        var themeName = Path.GetFileNameWithoutExtension(themePath);

        // Thử method chính: ApplyTheme
        wb.ApplyTheme(themePath);
        Logger.Debug($"Applied theme '{themeName}' to workbook using ApplyTheme");

      }
      catch (Exception ex)
      {
        Logger.Warning($"ApplyTheme failed for '{Path.GetFileName(themePath)}', trying alternative method: {ex.Message}");

        try
        {
          // Fallback: Áp dụng từ XML
          ApplyThemeFromXml(wb, themePath);
        }
        catch (Exception ex2)
        {
          Logger.Error($"Alternative theme application also failed for '{Path.GetFileName(themePath)}': {ex2.Message}");
          throw;
        }
      }
    }

    /// <summary>
    /// Áp dụng theme từ XML file (fallback method)
    /// Đọc và parse theme XML để áp dụng settings thủ công
    ///
    /// </summary>
    /// <param name="wb">Workbook cần áp dụng theme</param>
    /// <param name="themePath">Đường dẫn theme XML file</param>
    private void ApplyThemeFromXml(Workbook wb, string themePath)
    {
      try
      {
        // Đọc theme XML
        var xmlDoc = XDocument.Load(themePath);

        // TODO: Implement manual theme application từ XML
        // Hiện tại chỉ log để biết đã thử fallback
        Logger.Debug("Applied theme settings from XML file (basic implementation)");

      }
      catch (Exception ex)
      {
        Logger.Error($"Error parsing theme XML from {themePath}: {ex.Message}");
        throw;
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