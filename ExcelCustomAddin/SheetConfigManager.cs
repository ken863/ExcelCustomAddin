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
      public string ReferenceColumnHeader { get; set; }
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
    }

    /// <summary>
    /// Cấu hình logging
    /// </summary>
    public class LoggingConfig
    {
      public string LogDirectory { get; set; } = "";
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
        // Lấy đường dẫn của assembly hiện tại
        return Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
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
              ReferenceColumnHeader = sheetElement.Element("ReferenceColumnHeader")?.Value ?? "",
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
              ReferenceColumnHeader = sheetElement.Element("ReferenceColumnHeader")?.Value ?? "",
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
                    ReferenceColumnHeader = "参考 No.",
                    NumberFormat = "D2",
                    Description = "Sheet chung với prefix 共通_"
                },
                new SheetConfig
                {
                    Name = "テスト項目",
                    Prefix = "エビデンス_",
                    ReferenceColumnHeader = "参考 No.",
                    NumberFormat = "D2",
                    Description = "Sheet test items với prefix エビデンス_"
                }
            };

      _extendedSheets = new List<SheetConfig>();
      _generalConfig = new GeneralConfig();
      _loggingConfig = new LoggingConfig();
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
