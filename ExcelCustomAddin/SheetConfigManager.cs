using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace ExcelCustomAddin
{
  /// <summary>
  /// Class quản lý cấu hình sheet từ file XML
  /// </summary>
  public class SheetConfigManager
  {
    /// <summary>
    /// Thông tin cấu hình cho một sheet
    /// </summary>
    public class SheetConfig
    {
      public string Name { get; set; }
      public string Prefix { get; set; }
      public string NumberFormat { get; set; }
      public string Description { get; set; }
    }

    /// <summary>
    /// Cấu hình chung
    /// </summary>
    public class GeneralConfig
    {
      public int HeaderRowIndex { get; set; } = 2;
      public bool AutoFillCell { get; set; } = true;
      public bool EnableDebugLog { get; set; } = true;
      public int StartingNumber { get; set; } = 1;
      public string PageBreakColumnName { get; set; } = "AR";
      public string EvidenceFontName { get; set; } = "MS PGothic";
      public int PrintAreaLastRowIdx { get; set; } = 38;
      public double ColumnWidth { get; set; } = 2.38;
      public double RowHeight { get; set; } = 12.6;
      public int FontSize { get; set; } = 11;
      public string PageOrientation { get; set; } = "Landscape";
      public string PaperSize { get; set; } = "A4";
      public int Zoom { get; set; } = 100;
      public bool FitToPagesWide { get; set; } = false;
      public bool FitToPagesTall { get; set; } = false;
      public double LeftMargin { get; set; } = 0.75;
      public double RightMargin { get; set; } = 0.75;
      public double TopMargin { get; set; } = 1.0;
      public double BottomMargin { get; set; } = 0.75;
      public double HeaderMargin { get; set; } = 0.5;
      public double FooterMargin { get; set; } = 0.5;
      public bool CenterHorizontally { get; set; } = true;
      public int WindowZoom { get; set; } = 100;
      public string ViewMode { get; set; } = "PageBreakPreview";
    }

    /// <summary>
    /// Cấu hình logging
    /// </summary>
    public class LoggingConfig
    {
      public string LogDirectory { get; set; } = @"C:\ExcelCustomAddin";
      public bool EnableFileLogging { get; set; } = true;
      public bool EnableDebugOutput { get; set; } = true;
      public string LogLevel { get; set; } = "DEBUG";
      public string LogFileName { get; set; } = "ExcelAddin";
    }

    private static readonly string ConfigFileName = "SheetConfig.xml";
    private static string ConfigFilePath => Path.Combine(GetAddInDirectory(), ConfigFileName);

    private static List<SheetConfig> _specialSheets;
    private static List<SheetConfig> _extendedSheets;
    private static GeneralConfig _generalConfig;
    private static LoggingConfig _loggingConfig;

    /// <summary>
    /// Lấy thư mục chứa add-in
    /// </summary>
    /// <returns></returns>
    private static string GetAddInDirectory()
    {
      try
      {
        // Sử dụng thư mục cố định C:\ExcelCustomAddin
        string configDirectory = @"C:\ExcelCustomAddin";

        // Tạo thư mục nếu chưa tồn tại
        if (!Directory.Exists(configDirectory))
        {
          Directory.CreateDirectory(configDirectory);
        }

        return configDirectory;
      }
      catch
      {
        // Fallback về thư mục hiện tại nếu có lỗi
        return Environment.CurrentDirectory;
      }
    }

    /// <summary>
    /// Load cấu hình từ file XML
    /// </summary>
    public static void LoadConfiguration()
    {
      try
      {
        if (!File.Exists(ConfigFilePath))
        {
          CreateDefaultConfig();
        }

        var doc = XDocument.Load(ConfigFilePath);
        var root = doc.Root;

        // Load Special Sheets
        _specialSheets = new List<SheetConfig>();
        var specialSheetsElement = root.Element("SpecialSheets");
        if (specialSheetsElement != null)
        {
          foreach (var sheetElement in specialSheetsElement.Elements("Sheet"))
          {
            _specialSheets.Add(new SheetConfig
            {
              Name = sheetElement.Element("Name")?.Value ?? "",
              Prefix = sheetElement.Element("Prefix")?.Value ?? "",
              NumberFormat = sheetElement.Element("NumberFormat")?.Value ?? "D2",
              Description = sheetElement.Element("Description")?.Value ?? ""
            });
          }
        }

        // Load Extended Sheets
        _extendedSheets = new List<SheetConfig>();
        var extendedSheetsElement = root.Element("ExtendedSheets");
        if (extendedSheetsElement != null)
        {
          foreach (var sheetElement in extendedSheetsElement.Elements("Sheet"))
          {
            _extendedSheets.Add(new SheetConfig
            {
              Name = sheetElement.Element("Name")?.Value ?? "",
              Prefix = sheetElement.Element("Prefix")?.Value ?? "",
              NumberFormat = sheetElement.Element("NumberFormat")?.Value ?? "D2",
              Description = sheetElement.Element("Description")?.Value ?? ""
            });
          }
        }

        // Load General Settings
        _generalConfig = new GeneralConfig();
        var generalElement = root.Element("GeneralSettings");
        if (generalElement != null)
        {
          if (int.TryParse(generalElement.Element("HeaderRowIndex")?.Value, out int headerRow))
            _generalConfig.HeaderRowIndex = headerRow;

          if (bool.TryParse(generalElement.Element("AutoFillCell")?.Value, out bool autoFill))
            _generalConfig.AutoFillCell = autoFill;

          if (bool.TryParse(generalElement.Element("EnableDebugLog")?.Value, out bool debugLog))
            _generalConfig.EnableDebugLog = debugLog;

          if (int.TryParse(generalElement.Element("StartingNumber")?.Value, out int startNum))
            _generalConfig.StartingNumber = startNum;

          _generalConfig.PageBreakColumnName = generalElement.Element("PageBreakColumnName")?.Value ?? "AR";
          _generalConfig.EvidenceFontName = generalElement.Element("EvidenceFontName")?.Value ?? "MS PGothic";

          if (int.TryParse(generalElement.Element("PrintAreaLastRowIdx")?.Value, out int lastRow))
            _generalConfig.PrintAreaLastRowIdx = lastRow;

          if (double.TryParse(generalElement.Element("ColumnWidth")?.Value, out double colWidth))
            _generalConfig.ColumnWidth = colWidth;

          if (double.TryParse(generalElement.Element("RowHeight")?.Value, out double rowHeight))
            _generalConfig.RowHeight = rowHeight;

          if (int.TryParse(generalElement.Element("FontSize")?.Value, out int fontSize))
            _generalConfig.FontSize = fontSize;

          _generalConfig.PageOrientation = generalElement.Element("PageOrientation")?.Value ?? "Landscape";
          _generalConfig.PaperSize = generalElement.Element("PaperSize")?.Value ?? "A4";

          if (int.TryParse(generalElement.Element("Zoom")?.Value, out int zoom))
            _generalConfig.Zoom = zoom;

          if (bool.TryParse(generalElement.Element("FitToPagesWide")?.Value, out bool fitWide))
            _generalConfig.FitToPagesWide = fitWide;

          if (bool.TryParse(generalElement.Element("FitToPagesTall")?.Value, out bool fitTall))
            _generalConfig.FitToPagesTall = fitTall;

          if (double.TryParse(generalElement.Element("LeftMargin")?.Value, out double leftMargin))
            _generalConfig.LeftMargin = leftMargin;

          if (double.TryParse(generalElement.Element("RightMargin")?.Value, out double rightMargin))
            _generalConfig.RightMargin = rightMargin;

          if (double.TryParse(generalElement.Element("TopMargin")?.Value, out double topMargin))
            _generalConfig.TopMargin = topMargin;

          if (double.TryParse(generalElement.Element("BottomMargin")?.Value, out double bottomMargin))
            _generalConfig.BottomMargin = bottomMargin;

          if (double.TryParse(generalElement.Element("HeaderMargin")?.Value, out double headerMargin))
            _generalConfig.HeaderMargin = headerMargin;

          if (double.TryParse(generalElement.Element("FooterMargin")?.Value, out double footerMargin))
            _generalConfig.FooterMargin = footerMargin;

          if (bool.TryParse(generalElement.Element("CenterHorizontally")?.Value, out bool centerH))
            _generalConfig.CenterHorizontally = centerH;

          if (int.TryParse(generalElement.Element("WindowZoom")?.Value, out int windowZoom))
            _generalConfig.WindowZoom = windowZoom;
        }

        // Load Logging Settings
        _loggingConfig = new LoggingConfig();
        var loggingElement = root.Element("LoggingSettings");
        if (loggingElement != null)
        {
          _loggingConfig.LogDirectory = loggingElement.Element("LogDirectory")?.Value ?? "";

          if (bool.TryParse(loggingElement.Element("EnableFileLogging")?.Value, out bool fileLogging))
            _loggingConfig.EnableFileLogging = fileLogging;

          if (bool.TryParse(loggingElement.Element("EnableDebugOutput")?.Value, out bool debugOutput))
            _loggingConfig.EnableDebugOutput = debugOutput;

          _loggingConfig.LogLevel = loggingElement.Element("LogLevel")?.Value ?? "DEBUG";
          _loggingConfig.LogFileName = loggingElement.Element("LogFileName")?.Value ?? "ExcelAddin";
        }

        // Configure Logger with loaded settings
        if (_loggingConfig != null)
        {
          Logger.Configure(_loggingConfig.LogDirectory, _generalConfig.EnableDebugLog, _loggingConfig.LogFileName);
        }

        Logger.Info($"Loaded configuration: {_specialSheets.Count} special sheets, {_extendedSheets.Count} extended sheets");
      }
      catch (Exception ex)
      {
        Logger.Error($"Error loading configuration: {ex.Message}", ex);
        CreateDefaultConfig();
      }
    }

    /// <summary>
    /// Tạo file cấu hình mặc định nếu chưa tồn tại
    /// </summary>
    private static void CreateDefaultConfig()
    {
      _specialSheets = new List<SheetConfig>
            {
                new SheetConfig
                {
                    Name = "共通",
                    Prefix = "共通_",
                    NumberFormat = "D2",
                    Description = "Sheet chung với prefix 共通_"
                },
                new SheetConfig
                {
                    Name = "テスト項目",
                    Prefix = "エビデンス_",
                    NumberFormat = "D2",
                    Description = "Sheet test items với prefix エビデンス_"
                }
            };

      _extendedSheets = new List<SheetConfig>();
      _generalConfig = new GeneralConfig();
      _loggingConfig = new LoggingConfig();

      // Tạo file XML mặc định
      try
      {
        var doc = new XDocument(
          new XElement("SheetConfiguration",
            new XElement("SpecialSheets",
              from sheet in _specialSheets
              select new XElement("Sheet",
                new XElement("Name", sheet.Name),
                new XElement("Prefix", sheet.Prefix),
                new XElement("NumberFormat", sheet.NumberFormat),
                new XElement("Description", sheet.Description)
              )
            ),
            new XElement("ExtendedSheets"),
            new XElement("GeneralSettings",
              new XElement("HeaderRowIndex", _generalConfig.HeaderRowIndex),
              new XElement("AutoFillCell", _generalConfig.AutoFillCell),
              new XElement("EnableDebugLog", _generalConfig.EnableDebugLog),
              new XElement("StartingNumber", _generalConfig.StartingNumber)
            ),
            new XElement("LoggingSettings",
              new XElement("LogDirectory", _loggingConfig.LogDirectory),
              new XElement("EnableFileLogging", _loggingConfig.EnableFileLogging),
              new XElement("EnableDebugOutput", _loggingConfig.EnableDebugOutput),
              new XElement("LogLevel", _loggingConfig.LogLevel),
              new XElement("LogFileName", _loggingConfig.LogFileName)
            )
          )
        );

        doc.Save(ConfigFilePath);
      }
      catch (Exception ex)
      {
        // Nếu không thể tạo file, chỉ log lỗi nhưng vẫn tiếp tục với config trong memory
        System.Diagnostics.Debug.WriteLine($"Could not create default config file: {ex.Message}");
      }
    }

    /// <summary>
    /// Lấy cấu hình cho sheet theo tên
    /// </summary>
    /// <param name="sheetName">Tên sheet</param>
    /// <returns>Cấu hình sheet hoặc null nếu không tìm thấy</returns>
    public static SheetConfig GetSheetConfig(string sheetName)
    {
      if (_specialSheets == null)
        LoadConfiguration();

      // Tìm trong special sheets trước
      var config = _specialSheets?.FirstOrDefault(s => s.Name == sheetName);
      if (config != null)
        return config;

      // Tìm trong extended sheets
      return _extendedSheets?.FirstOrDefault(s => s.Name == sheetName);
    }

    /// <summary>
    /// Lấy cấu hình chung
    /// </summary>
    /// <returns>Cấu hình chung</returns>
    public static GeneralConfig GetGeneralConfig()
    {
      if (_generalConfig == null)
        LoadConfiguration();

      return _generalConfig ?? new GeneralConfig();
    }

    /// <summary>
    /// Lấy tất cả sheet configs
    /// </summary>
    /// <returns>Danh sách tất cả các sheet config</returns>
    public static List<SheetConfig> GetAllSheetConfigs()
    {
      if (_specialSheets == null)
        LoadConfiguration();

      var allConfigs = new List<SheetConfig>();
      if (_specialSheets != null)
        allConfigs.AddRange(_specialSheets);
      if (_extendedSheets != null)
        allConfigs.AddRange(_extendedSheets);

      return allConfigs;
    }

    /// <summary>
    /// Reload cấu hình từ file
    /// </summary>
    public static void ReloadConfiguration()
    {
      _specialSheets = null;
      _extendedSheets = null;
      _generalConfig = null;
      _loggingConfig = null;
      LoadConfiguration();
    }

    /// <summary>
    /// Lấy cấu hình logging
    /// </summary>
    /// <returns>LoggingConfig instance</returns>
    public static LoggingConfig GetLoggingConfig()
    {
      if (_loggingConfig == null)
      {
        LoadConfiguration();
      }
      return _loggingConfig ?? new LoggingConfig();
    }

    /// <summary>
    /// Kiểm tra xem có sheet config nào không
    /// </summary>
    /// <param name="sheetName">Tên sheet</param>
    /// <returns>True nếu có cấu hình cho sheet này</returns>
    public static bool HasSheetConfig(string sheetName)
    {
      return GetSheetConfig(sheetName) != null;
    }
  }
}
