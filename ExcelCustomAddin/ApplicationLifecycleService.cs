using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using System;
using System.Linq;
using System.Windows.Threading;

namespace ExcelCustomAddin
{
  /// <summary>
  /// Service xử lý các sự kiện lifecycle và quản lý action panel của application
  /// </summary>
  public class ApplicationLifecycleService
  {
    private readonly ThisAddIn _addIn;

    public ApplicationLifecycleService(ThisAddIn addIn)
    {
      _addIn = addIn;
    }

    /// <summary>
    /// ThisAddIn_Startup - Khởi động add-in
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

      // Register Hanle Events
      ((AppEvents_Event)Globals.ThisAddIn.Application).NewWorkbook += Application_NewWorkbook;
      Globals.ThisAddIn.Application.WorkbookOpen += Application_WorkbookOpen;
      Globals.ThisAddIn.Application.WorkbookActivate += Application_WorkbookActive;
      Globals.ThisAddIn.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
      Globals.ThisAddIn.Application.WorkbookAfterSave += Application_WorkbookAfterSave;
      Globals.ThisAddIn.Application.SheetActivate += Application_SheetActivate;

      // Tạo ActionPane cho workbook hiện tại (nếu có) với delay để tránh trùng lặp
      if (_addIn.Application.ActiveWorkbook != null)
      {
        // Sử dụng timer để đảm bảo chỉ tạo 1 lần sau khi startup xong
        var timer = new System.Windows.Forms.Timer();
        timer.Interval = 500; // 500ms delay
        timer.Tick += (s, args) =>
        {
          timer.Stop();
          timer.Dispose();
          CreateActionsPane(_addIn.Application.ActiveWorkbook);
        };
        timer.Start();
      }
    }

    /// <summary>
    /// ThisAddIn_Shutdown - Cleanup events
    /// </summary>
    public void Shutdown()
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
        if (_addIn._actionPanel != null)
        {
          _addIn._actionPanel.CreateEvidenceEvent -= _addIn.CreateEvidence;
          _addIn._actionPanel.FormatImagesEvent -= _addIn.FormatImages;
          _addIn._actionPanel.FormatDocumentEvent -= _addIn.FormatDocument;
          _addIn._actionPanel.ChangeSheetNameEvent -= _addIn.ChangeSheetName;
          _addIn._actionPanel.InsertMultipleImagesEvent -= _addIn.InsertMultipleImages;
          _addIn._actionPanel.PinSheetEvent -= _addIn.PinSheet;
          _addIn._actionPanel.listofSheet.SelectedIndexChanged -= _addIn._sheetService.ListOfSheet_SelectionChanged;
        }
      }
      catch (Exception ex)
      {
        // Log error if needed, but don't show MessageBox during shutdown
        Logger.Error($"Error during shutdown: {ex.Message}", ex);
      }
    }

    /// <summary>
    /// Application_NewWorkbook
    /// </summary>
    private void Application_NewWorkbook(Workbook Wb)
    {
      Logger.Debug($"Application_NewWorkbook called for: {Wb?.Name}");
      LoadTemplate(Wb);
      CreateActionsPane(Wb);
    }

    /// <summary>
    /// Application_WorkbookOpen
    /// </summary>
    private void Application_WorkbookOpen(Workbook Wb)
    {
      Logger.Debug($"Application_WorkbookOpen called for: {Wb?.Name}");
      LoadTemplate(Wb);
      CreateActionsPane(Wb);
    }

    /// <summary>
    /// Application_WorkbookActivate
    /// </summary>
    private void Application_WorkbookActive(Workbook Wb)
    {
      // Khi activate workbook, chỉ cập nhật action panel nếu đã tồn tại
      // Không tạo mới để tránh trùng lặp với Open/New events
      if (Wb != null && ThisAddIn.CreatedActionPanes.Contains(Wb.Name))
      {
        // Chỉ cập nhật nếu đã có action panel
        if (_addIn._actionPanel != null)
        {
          var currentSheetName = Wb.ActiveSheet?.Name;
          _addIn._actionPanel.BindSheetList(_addIn._sheetService.GetListOfSheet(), currentSheetName);
        }
      }
    }

    /// <summary>
    /// Application_WorkbookBeforeClose
    /// </summary>
    private void Application_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
    {
      if (Wb != null)
      {
        string workbookKey = Wb.Name;

        // Xóa workbook khỏi danh sách đã tạo action panel
        if (ThisAddIn.CreatedActionPanes.Contains(workbookKey))
        {
          ThisAddIn.CreatedActionPanes.Remove(workbookKey);
        }

        // Xóa pinned sheets của workbook này
        if (ThisAddIn.PinnedSheets.ContainsKey(workbookKey))
        {
          ThisAddIn.PinnedSheets.Remove(workbookKey);
        }
      }
    }

    /// <summary>
    /// Application_WorkbookAfterSave - Cập nhật file path sau khi lưu
    /// </summary>
    private void Application_WorkbookAfterSave(Workbook Wb, bool Success)
    {
      if (Wb != null && Success && _addIn._actionPanel != null)
      {
        // Cập nhật file path display sau khi workbook được lưu thành công
        _addIn._actionPanel.RefreshFilePathDisplay();
        Logger.Debug($"File path refreshed after save for: {Wb.Name}");
      }
    }

    /// <summary>
    /// Application_SheetActivate - Xử lý khi chuyển đổi sheet
    /// </summary>
    private void Application_SheetActivate(object Sh)
    {
      try
      {
        var worksheet = Sh as Worksheet;
        if (worksheet != null)
        {
          Logger.Debug($"Sheet activated: {worksheet.Name}");

          // Đánh dấu sheet đang được activate để tránh sự kiện update danh sách sheet
          _addIn.IsSheetActivating = true;

          // Cập nhật danh sách sheet trong action panel
          if (_addIn._actionPanel != null)
          {
            _addIn._actionPanel.BindSheetList(_addIn._sheetService.GetListOfSheet(), worksheet.Name);
          }

          // Reset flag sau một khoảng thời gian ngắn
          var timer = new System.Windows.Forms.Timer();
          timer.Interval = 100; // 100ms
          timer.Tick += (s, args) =>
          {
            timer.Stop();
            timer.Dispose();
            _addIn.IsSheetActivating = false;
          };
          timer.Start();
        }
      }
      catch (Exception ex)
      {
        Logger.Error($"Error in Application_SheetActivate: {ex.Message}", ex);
      }
    }

    /// <summary>
    /// Load template cho workbook mới
    /// Public method để load template
    /// </summary>
    public void LoadTemplate(Workbook Wb)
    {
      if (Wb != null)
      {
        try
        {
          // Có thể thêm logic load template ở đây nếu cần
          Logger.Debug($"Template loaded for workbook: {Wb.Name}");
        }
        catch (Exception ex)
        {
          Logger.Error($"Error loading template for workbook {Wb.Name}: {ex.Message}", ex);
        }
      }
    }

    /// <summary>
    /// CreateActionsPane - Tạo action panel cho workbook
    /// </summary>
    private void CreateActionsPane(Workbook Wb)
    {
      if (Wb != null)
      {
        string workbookKey = Wb.Name;

        lock (ThisAddIn._lockObject)
        {
          // Debug logging
          Logger.Debug($"CreateActionsPane called for: {workbookKey}");

          // Kiểm tra xem action panel đã được tạo cho workbook này chưa
          if (ThisAddIn.CreatedActionPanes.Contains(workbookKey))
          {
            Logger.Debug($"Action panel already exists for: {workbookKey}, updating only");
            // Nếu đã tạo rồi, chỉ cần cập nhật danh sách sheet
            if (_addIn._actionPanel != null && _addIn.myCustomTaskPane != null)
            {
              // Đảm bảo task pane đang active cho workbook này
              var currentTaskPane = TaskPaneManager.GetTaskPane(workbookKey, "WORKSHEET TOOLS", null);
              if (currentTaskPane != null)
              {
                _addIn.myCustomTaskPane = currentTaskPane;
                _addIn._actionPanel = (ActionPanelControl)_addIn.myCustomTaskPane.Control;

                var currentSheetName = Wb.ActiveSheet?.Name;
                _addIn._actionPanel.BindSheetList(_addIn._sheetService.GetListOfSheet(), currentSheetName);
              }
            }
            return;
          }

          Logger.Debug($"Creating new action panel for: {workbookKey}");

          // Get Active ActionsPanel
          _addIn.myCustomTaskPane = TaskPaneManager.GetTaskPane(Wb.Name, "WORKSHEET TOOLS", () => new ActionPanelControl());
          _addIn._actionPanel = (ActionPanelControl)_addIn.myCustomTaskPane?.Control;

          if (_addIn._actionPanel != null)
          {
            // Hủy đăng ký các event cũ trước khi đăng ký mới để tránh đăng ký trùng lặp
            _addIn._actionPanel.CreateEvidenceEvent -= _addIn.CreateEvidence;
            _addIn._actionPanel.FormatImagesEvent -= _addIn.FormatImages;
            _addIn._actionPanel.FormatDocumentEvent -= _addIn.FormatDocument;
            _addIn._actionPanel.ChangeSheetNameEvent -= _addIn.ChangeSheetName;
            _addIn._actionPanel.InsertMultipleImagesEvent -= _addIn.InsertMultipleImages;
            _addIn._actionPanel.PinSheetEvent -= _addIn.PinSheet;
            _addIn._actionPanel.listofSheet.SelectedIndexChanged -= _addIn._sheetService.ListOfSheet_SelectionChanged;

            // Đăng ký các event mới
            _addIn._actionPanel.CreateEvidenceEvent += _addIn.CreateEvidence;
            _addIn._actionPanel.FormatImagesEvent += _addIn.FormatImages;
            _addIn._actionPanel.FormatDocumentEvent += _addIn.FormatDocument;
            _addIn._actionPanel.ChangeSheetNameEvent += _addIn.ChangeSheetName;
            _addIn._actionPanel.InsertMultipleImagesEvent += _addIn.InsertMultipleImages;
            _addIn._actionPanel.PinSheetEvent += _addIn.PinSheet;
            _addIn._actionPanel.listofSheet.SelectedIndexChanged += _addIn._sheetService.ListOfSheet_SelectionChanged;

            // Cập nhật danh sách sheet và chọn sheet hiện tại khi tạo ActionPane
            var currentSheetName = Wb.ActiveSheet?.Name;
            _addIn._actionPanel.BindSheetList(_addIn._sheetService.GetListOfSheet(), currentSheetName);

            // *** THÊM DÒNG NÀY: Tự động hiển thị Action Panel khi workbook được mở ***
            _addIn.myCustomTaskPane.Visible = true;

            // Tùy chọn: Đặt độ rộng mặc định cho task pane (tuỳ chỉnh theo nhu cầu)
            _addIn.myCustomTaskPane.Width = 300;

            // Đánh dấu workbook này đã được tạo action panel
            ThisAddIn.CreatedActionPanes.Add(workbookKey);
            Logger.Debug($"Action panel created and marked for: {workbookKey}");
          }
        }
      }
    }
  }
}