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
  /// Service xử lý các sự kiện lifecycle và quản lý action panel của application
  /// Hỗ trợ nhiều version Excel và architecture x64/x86
  /// </summary>
  public class ApplicationLifecycleService
  {
    #region Fields

    private readonly ThisAddIn _addIn;

    // Excel version và architecture info
    private readonly string _excelVersion;
    private readonly string _architecture;
    private readonly bool _isExcel64Bit;

    #endregion

    #region Constructor

    public ApplicationLifecycleService(ThisAddIn addIn)
    {
      _addIn = addIn;

      // Detect Excel version và architecture
      _excelVersion = GetExcelVersion();
      _architecture = GetSystemArchitecture();
      _isExcel64Bit = IsExcel64Bit();

      Logger.Info($"Excel Version: {_excelVersion}, Architecture: {_architecture}, Excel x64: {_isExcel64Bit}");
    }

    #endregion

    #region Public Methods

    /// <summary>
    /// Khởi động add-in và đăng ký các events cần thiết
    /// </summary>
    public void Startup()
    {
      Logger.Info("Excel Custom Add-in starting up...");

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
      _addIn._dispatcher = Dispatcher.CurrentDispatcher;

      // Register application events
      RegisterApplicationEvents();

      // Tạo ActionPane cho workbook hiện tại (nếu có) với delay để tránh trùng lặp
      if (_addIn.Application.ActiveWorkbook != null)
      {
        // Sử dụng timer để đảm bảo chỉ tạo 1 lần sau khi startup xong
        var timer = new System.Windows.Forms.Timer { Interval = 500 };
        timer.Tick += (s, args) => { timer.Stop(); timer.Dispose(); CreateActionsPane(_addIn.Application.ActiveWorkbook); };
        timer.Start();
      }
    }

    /// <summary>
    /// Cleanup events khi shutdown add-in
    /// </summary>
    public void Shutdown()
    {
      try
      {
        // Hủy đăng ký các application events
        UnregisterApplicationEvents();

        // Hủy đăng ký action panel events
        UnregisterActionPanelEvents();
      }
      catch (Exception ex)
      {
        Logger.Error($"Error during shutdown: {ex.Message}", ex);
      }
    }

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

        // Xác định đường dẫn cơ sở dựa trên phiên bản Office và architecture
        string[] possiblePaths = GetPossibleOfficePaths(officeVersion);

        foreach (var path in possiblePaths)
        {
          if (System.IO.Directory.Exists(path))
          {
            officeBasePath = path;
            Logger.Debug($"Found Office themes path: {officeBasePath}");
            break;
          }
        }

        if (string.IsNullOrEmpty(officeBasePath))
        {
          Logger.Error($"Không tìm thấy thư mục Document Themes cho phiên bản Office {officeVersion}. Vui lòng kiểm tra cài đặt Office.");
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
    /// Lấy version Excel hiện tại
    /// </summary>
    private string GetExcelVersion()
    {
      try
      {
        return $"Excel {Globals.ThisAddIn.Application.Version}";
      }
      catch
      {
        return "Unknown";
      }
    }

    /// <summary>
    /// Lấy architecture của hệ thống
    /// </summary>
    private string GetSystemArchitecture() => Environment.Is64BitProcess ? "x64" : "x86";

    /// <summary>
    /// Kiểm tra Excel có phải x64 không
    /// </summary>
    private bool IsExcel64Bit()
    {
      try
      {
        return Marshal.SizeOf(IntPtr.Zero) == 8;
      }
      catch
      {
        return Environment.Is64BitProcess;
      }
    }

    #endregion

    #region Office Path Detection

    /// <summary>
    /// Lấy danh sách các đường dẫn có thể có cho Office themes dựa trên phiên bản
    /// </summary>
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
          return new string[0];
      }

      // Lấy các đường dẫn Program Files có thể có
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
            possiblePaths.Add(System.IO.Path.Combine(programFiles, "Microsoft Office", "Office15", themeFolder));
          }
          // Office 2016+: root\Document Themes 16
          else if (officeVersion == "16.0")
          {
            possiblePaths.Add(System.IO.Path.Combine(programFiles, "Microsoft Office", "root", themeFolder));
          }
        }
      }

      return possiblePaths.ToArray();
    }

    #endregion

    #region Event Registration

    /// <summary>
    /// Register tất cả application events
    /// </summary>
    private void RegisterApplicationEvents()
    {
      var app = Globals.ThisAddIn.Application;
      ((AppEvents_Event)app).NewWorkbook += Application_NewWorkbook;
      app.WorkbookOpen += Application_WorkbookOpen;
      app.WorkbookActivate += Application_WorkbookActive;
      app.WorkbookBeforeClose += Application_WorkbookBeforeClose;
      app.WorkbookAfterSave += Application_WorkbookAfterSave;
      app.SheetActivate += Application_SheetActivate;
    }

    /// <summary>
    /// Unregister tất cả application events
    /// </summary>
    private void UnregisterApplicationEvents()
    {
      var app = Globals.ThisAddIn.Application;
      if (app != null)
      {
        ((AppEvents_Event)app).NewWorkbook -= Application_NewWorkbook;
        app.WorkbookOpen -= Application_WorkbookOpen;
        app.WorkbookActivate -= Application_WorkbookActive;
        app.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
        app.WorkbookAfterSave -= Application_WorkbookAfterSave;
        app.SheetActivate -= Application_SheetActivate;
      }
    }

    /// <summary>
    /// Đăng ký events cho ActionPanel
    /// </summary>
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
    /// Hủy đăng ký events cho ActionPanel
    /// </summary>
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
    /// Unregister action panel events (overload)
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
    /// </summary>
    private void Application_NewWorkbook(Workbook Wb) => HandleWorkbookOpened(Wb, "NewWorkbook");

    /// <summary>
    /// Xử lý sự kiện mở workbook
    /// </summary>
    private void Application_WorkbookOpen(Workbook Wb) => HandleWorkbookOpened(Wb, "WorkbookOpen");

    /// <summary>
    /// Handle workbook opened events (New/Open)
    /// </summary>
    private void HandleWorkbookOpened(Workbook Wb, string eventName)
    {
      Logger.Debug($"Application_{eventName} called for: {Wb?.Name}");
      LoadTemplate(Wb);
      CreateActionsPane(Wb);
    }

    /// <summary>
    /// Xử lý sự kiện activate workbook
    /// </summary>
    private void Application_WorkbookActive(Workbook Wb)
    {
      if (Wb != null && ThisAddIn.CreatedActionPanes.Contains(Wb.Name))
        UpdateActionPanelForWorkbook(Wb, Wb.Name);
    }

    /// <summary>
    /// Xử lý sự kiện trước khi đóng workbook
    /// </summary>
    private void Application_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
    {
      if (Wb != null)
        CleanupWorkbookResources(Wb.Name);
    }

    /// <summary>
    /// Xử lý sự kiện sau khi lưu workbook
    /// </summary>
    private void Application_WorkbookAfterSave(Workbook Wb, bool Success)
    {
      if (Wb != null && Success && _addIn._actionPanel != null)
      {
        _addIn._actionPanel.RefreshFilePathDisplay();
        Logger.Debug($"File path refreshed after save for: {Wb.Name}");
      }
    }

    /// <summary>
    /// Xử lý sự kiện activate sheet
    /// </summary>
    private void Application_SheetActivate(object Sh)
    {
      try
      {
        var worksheet = Sh as Worksheet;
        if (worksheet == null) return;

        Logger.Debug($"Sheet activated: {worksheet.Name}");

        // Đánh dấu sheet đang được activate để tránh sự kiện update danh sách sheet
        _addIn.IsSheetActivating = true;

        // Cập nhật danh sách sheet trong action panel
        _addIn._actionPanel?.BindSheetList(_addIn._sheetService.GetListOfSheet(), worksheet.Name);

        // Reset flag sau một khoảng thời gian ngắn
        var timer = new System.Windows.Forms.Timer { Interval = 100 };
        timer.Tick += (s, args) => { timer.Stop(); timer.Dispose(); _addIn.IsSheetActivating = false; };
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
    /// Tạo action panel cho workbook
    /// </summary>
    private void CreateActionsPane(Workbook Wb)
    {
      if (Wb == null) return;

      string workbookKey = Wb.Name;

      lock (ThisAddIn._lockObject)
      {
        Logger.Debug($"CreateActionsPane called for: {workbookKey}");

        if (ThisAddIn.CreatedActionPanes.Contains(workbookKey))
        {
          Logger.Debug($"Action panel already exists for: {workbookKey}, updating only");
          UpdateActionPanelForWorkbook(Wb, workbookKey);
          return;
        }

        Logger.Debug($"Creating new action panel for: {workbookKey}");
        CreateNewActionPanel(Wb, workbookKey);
      }
    }

    /// <summary>
    /// Cập nhật ActionPanel cho workbook hiện tại
    /// </summary>
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
    /// Tạo mới ActionPanel cho workbook
    /// </summary>
    private void CreateNewActionPanel(Workbook wb, string workbookKey)
    {
      _addIn.myCustomTaskPane = TaskPaneManager.GetTaskPane(wb.Name, "WORKSHEET TOOLS", () => new ActionPanelControl());
      _addIn._actionPanel = (ActionPanelControl)_addIn.myCustomTaskPane?.Control;

      if (_addIn._actionPanel != null)
      {
        UnregisterActionPanelEvents(_addIn._actionPanel);
        RegisterActionPanelEvents(_addIn._actionPanel);

        var currentSheetName = wb.ActiveSheet?.Name;
        _addIn._actionPanel.BindSheetList(_addIn._sheetService.GetListOfSheet(), currentSheetName);

        _addIn.myCustomTaskPane.Visible = true;
        _addIn.myCustomTaskPane.Width = 300;

        ThisAddIn.CreatedActionPanes.Add(workbookKey);
        Logger.Debug($"Action panel created and marked for: {workbookKey}");
      }
    }

    /// <summary>
    /// Cleanup resources cho workbook khi đóng
    /// </summary>
    private void CleanupWorkbookResources(string workbookKey)
    {
      ThisAddIn.CreatedActionPanes.Remove(workbookKey);
      ThisAddIn.PinnedSheets.Remove(workbookKey);
    }

    #endregion
  }
}