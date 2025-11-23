using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace ExcelCustomAddin
{
    /// <summary>
    /// SheetConfigManager - Quản lý cấu hình từ file XML
    ///
    /// Chức năng chính:
    /// - Load và parse cấu hình từ SheetConfig.xml
    /// - Quản lý special sheets và extended sheets
    /// - Cung cấp general settings và logging config
    /// - Tự động tạo default config nếu file không tồn tại
    /// - Thread-safe lazy loading với static initialization
    ///
    /// Cấu trúc file XML:
    /// - SpecialSheets: Sheets đặc biệt (共通, テスト項目)
    /// - ExtendedSheets: Sheets mở rộng (tùy chỉnh)
    /// - GeneralSettings: Cài đặt chung (font, margins, zoom, etc.)
    /// - LoggingSettings: Cấu hình logging
    ///
    /// Tác giả: lam.pt
    /// Ngày tạo: 2025
    /// </summary>
    public class SheetConfigManager
    {
        #region Configuration Classes

        /// <summary>
        /// Thông tin cấu hình cho một sheet
        /// Định nghĩa prefix, format số, và mô tả cho từng loại sheet
        ///
        /// </summary>
        public class SheetConfig
        {
            /// <summary>Tên sheet (VD: "共通", "テスト項目")</summary>
            public string Name { get; set; }

            /// <summary>Prefix để generate tên sheet (VD: "共通_", "エビデンス_")</summary>
            public string Prefix { get; set; }

            /// <summary>Reference column header cho sheet</summary>
            public string ReferenceColumnHeader { get; set; }

            /// <summary>Format số cho auto-numbering (VD: "D2" = 01, 02, ...)</summary>
            public string NumberFormat { get; set; }

            /// <summary>Mô tả chức năng của sheet config</summary>
            public string Description { get; set; }

            public string IsHorizontal { get; set; }
        }

        /// <summary>
        /// Cấu hình chung cho toàn bộ add-in
        /// Bao gồm page setup, fonts, margins, zoom settings
        ///
        /// </summary>
        public class GeneralConfig
        {
            // Sheet processing settings
            public int HeaderRowIndex { get; set; } = 2;
            public bool AutoFillCell { get; set; } = true;
            public bool EnableDebugLog { get; set; } = true;
            public int StartingNumber { get; set; } = 1;

            // Page layout settings
            public string PageBreakColumnName { get; set; } = "AR";
            public string EvidenceFontName { get; set; } = "MS PGothic";
            public string BackButtonFontName { get; set; } = "Calibri";
            public int PrintAreaLastRowIdx { get; set; } = 38;

            // Cell formatting
            public double ColumnWidth { get; set; } = 2.38;
            public double RowHeight { get; set; } = 12.6;
            public int FontSize { get; set; } = 11;

            // Page setup
            public string PageOrientation { get; set; } = "Landscape";
            public string PaperSize { get; set; } = "A4";
            public int Zoom { get; set; } = 100;
            public bool FitToPagesWide { get; set; } = false;
            public bool FitToPagesTall { get; set; } = false;

            // Margins
            public double LeftMargin { get; set; } = 0.75;
            public double RightMargin { get; set; } = 0.75;
            public double TopMargin { get; set; } = 1.0;
            public double BottomMargin { get; set; } = 1.0;
            public double HeaderMargin { get; set; } = 0.5;
            public double FooterMargin { get; set; } = 0.5;
            public bool CenterHorizontally { get; set; } = true;

            // View settings
            public int WindowZoom { get; set; } = 100;
            public string ViewMode { get; set; } = "PageBreakPreview";
        }

        /// <summary>
        /// Cấu hình logging cho add-in
        /// Điều khiển file logging, debug output, và log levels
        ///
        /// </summary>
        public class LoggingConfig
        {
            public string LogDirectory { get; set; } = @"C:\ExcelCustomAddin";
            public bool EnableFileLogging { get; set; } = true;
            public bool EnableDebugOutput { get; set; } = true;
            public string LogLevel { get; set; } = "DEBUG";
            public string LogFileName { get; set; } = "ExcelAddin";
        }

        #endregion

        #region Static Fields

        private static readonly string ConfigFileName = "SheetConfig.xml";
        private static string ConfigFilePath => Path.Combine(GetAddInDirectory(), ConfigFileName);

        // Cached configuration instances
        internal static List<SheetConfig> _sheets;
        internal static GeneralConfig _generalConfig;
        internal static LoggingConfig _loggingConfig;

        #endregion

        #region Directory Management

        /// <summary>
        /// Lấy thư mục chứa add-in và config files
        /// Sử dụng thư mục cố định C:\ExcelCustomAddin cho consistency
        ///
        /// </summary>
        /// <returns>Đường dẫn thư mục config</returns>
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
        /// Lấy đường dẫn file config nguồn từ thư mục gốc của add-in
        /// Tìm kiếm trong các vị trí có thể có của file config
        ///
        /// </summary>
        /// <returns>Đường dẫn file config nguồn hoặc null nếu không tìm thấy</returns>
        private static string GetSourceConfigPath()
        {
            try
            {
                // Lấy đường dẫn của assembly hiện tại
                string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string assemblyDir = Path.GetDirectoryName(assemblyPath);

                // File config nguồn có thể ở cùng thư mục với assembly
                string sourcePath = Path.Combine(assemblyDir, ConfigFileName);
                if (File.Exists(sourcePath))
                {
                    return sourcePath;
                }

                // Có thể ở thư mục cha (bin/Debug hoặc bin/Release)
                string parentDir = Directory.GetParent(assemblyDir)?.FullName;
                if (parentDir != null)
                {
                    sourcePath = Path.Combine(parentDir, ConfigFileName);
                    if (File.Exists(sourcePath))
                    {
                        return sourcePath;
                    }
                }

                // Có thể ở thư mục gốc của project
                string projectRoot = Directory.GetParent(parentDir ?? assemblyDir)?.FullName;
                if (projectRoot != null)
                {
                    sourcePath = Path.Combine(projectRoot, ConfigFileName);
                    if (File.Exists(sourcePath))
                    {
                        return sourcePath;
                    }
                }

                return null; // Không tìm thấy file source
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error finding source config path: {ex.Message}");
                return null;
            }
        }

        #endregion

        #region Configuration Loading

        /// <summary>
        /// Load cấu hình từ file XML
        /// Main entry point cho việc load configuration
        ///
        /// Quy trình:
        /// 1. Tìm file config trong thư mục đích
        /// 2. Copy từ source nếu cần thiết
        /// 3. Parse XML và load vào memory
        /// 4. Fallback to default config nếu thất bại
        ///
        /// </summary>
        public static void LoadConfiguration()
        {
            try
            {
                string configDir = GetAddInDirectory();
                string configPath = Path.Combine(configDir, ConfigFileName);

                // Kiểm tra xem file config có tồn tại trong thư mục config không
                if (!File.Exists(configPath))
                {
                    // Thử copy từ thư mục gốc của add-in nếu có
                    string sourceConfigPath = GetSourceConfigPath();
                    if (!string.IsNullOrEmpty(sourceConfigPath) && File.Exists(sourceConfigPath))
                    {
                        try
                        {
                            // Tạo thư mục nếu chưa tồn tại
                            Directory.CreateDirectory(configDir);

                            // Copy file config từ source
                            File.Copy(sourceConfigPath, configPath, true);
                            Logger.Info($"Copied config file from {sourceConfigPath} to {configPath}");
                        }
                        catch (Exception copyEx)
                        {
                            Logger.Warning($"Could not copy config file from {sourceConfigPath}: {copyEx.Message}");
                            // Fallback to create default config
                            CreateDefaultConfig();
                            return;
                        }
                    }
                    else
                    {
                        // Không có file source, tạo default config
                        CreateDefaultConfig();
                        return;
                    }
                }

                // Load config từ file
                LoadConfigFromFile(configPath);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error loading configuration: {ex.Message}", ex);
                CreateDefaultConfig();
            }
        }

        /// <summary>
        /// Load cấu hình từ file XML đã tồn tại
        /// Parse XML và populate các config objects
        ///
        /// </summary>
        /// <param name="configPath">Đường dẫn file config</param>
        private static void LoadConfigFromFile(string configPath)
        {
            var doc = XDocument.Load(configPath);
            var root = doc.Root;

            // Load Sheets
            _sheets = new List<SheetConfig>();
            var sheetsElement = root.Element("Sheets");
            
            if (sheetsElement != null)
            {
                // New unified structure
                foreach (var sheetElement in sheetsElement.Elements("Sheet"))
                {
                    var sheetConfig = new SheetConfig
                    {
                        Name = sheetElement.Element("Name")?.Value ?? "",
                        Prefix = sheetElement.Element("Prefix")?.Value ?? "",
                        ReferenceColumnHeader = sheetElement.Element("ReferenceColumnHeader")?.Value ?? "",
                        NumberFormat = sheetElement.Element("NumberFormat")?.Value ?? "D2",
                        Description = sheetElement.Element("Description")?.Value ?? "",
                        IsHorizontal = sheetElement.Element("IsHorizontal")?.Value ?? "false",
                    };
                    _sheets.Add(sheetConfig);
                    Logger.Debug($"Loaded sheet: {sheetConfig.Name}, Prefix: {sheetConfig.Prefix}, RefHeader: {sheetConfig.ReferenceColumnHeader}");
                }
                Logger.Info($"Loaded {_sheets.Count} sheets from unified <Sheets> element");
            }
            else
            {
                // Fallback to old structure
                Logger.Info("Unified <Sheets> element not found, trying legacy structure...");
                
                // Load Special Sheets
                var specialSheetsElement = root.Element("SpecialSheets");
                if (specialSheetsElement != null)
                {
                    foreach (var sheetElement in specialSheetsElement.Elements("Sheet"))
                    {
                        var sheetConfig = new SheetConfig
                        {
                            Name = sheetElement.Element("Name")?.Value ?? "",
                            Prefix = sheetElement.Element("Prefix")?.Value ?? "",
                            ReferenceColumnHeader = sheetElement.Element("ReferenceColumnHeader")?.Value ?? "",
                            NumberFormat = sheetElement.Element("NumberFormat")?.Value ?? "D2",
                            Description = sheetElement.Element("Description")?.Value ?? "",
                            IsHorizontal = sheetElement.Element("IsHorizontal")?.Value ?? "false",
                        };
                        _sheets.Add(sheetConfig);
                        Logger.Debug($"Loaded special sheet: {sheetConfig.Name}, Prefix: {sheetConfig.Prefix}");
                    }
                }

                // Load Extended Sheets
                var extendedSheetsElement = root.Element("ExtendedSheets");
                if (extendedSheetsElement != null)
                {
                    foreach (var sheetElement in extendedSheetsElement.Elements("Sheet"))
                    {
                        var sheetConfig = new SheetConfig
                        {
                            Name = sheetElement.Element("Name")?.Value ?? "",
                            Prefix = sheetElement.Element("Prefix")?.Value ?? "",
                            ReferenceColumnHeader = sheetElement.Element("ReferenceColumnHeader")?.Value ?? "",
                            NumberFormat = sheetElement.Element("NumberFormat")?.Value ?? "D2",
                            Description = sheetElement.Element("Description")?.Value ?? "",
                            IsHorizontal = sheetElement.Element("IsHorizontal")?.Value ?? "false",
                        };
                        _sheets.Add(sheetConfig);
                        Logger.Debug($"Loaded extended sheet: {sheetConfig.Name}, Prefix: {sheetConfig.Prefix}");
                    }
                }
                
                if (_sheets.Count > 0)
                {
                    Logger.Info($"Loaded {_sheets.Count} sheets from legacy structure (SpecialSheets + ExtendedSheets)");
                }
                else
                {
                    Logger.Warning("No sheets found in either unified or legacy XML structure");
                }
            }

            // Load General Settings
            _generalConfig = new GeneralConfig();
            var generalElement = root.Element("GeneralSettings");
            if (generalElement != null)
            {
                // Parse numeric values with TryParse for safety
                if (int.TryParse(generalElement.Element("HeaderRowIndex")?.Value, out int headerRow))
                    _generalConfig.HeaderRowIndex = headerRow;

                if (bool.TryParse(generalElement.Element("AutoFillCell")?.Value, out bool autoFill))
                    _generalConfig.AutoFillCell = autoFill;

                if (bool.TryParse(generalElement.Element("EnableDebugLog")?.Value, out bool debugLog))
                    _generalConfig.EnableDebugLog = debugLog;

                if (int.TryParse(generalElement.Element("StartingNumber")?.Value, out int startNum))
                    _generalConfig.StartingNumber = startNum;

                // Parse string values with null coalescing
                _generalConfig.PageBreakColumnName = generalElement.Element("PageBreakColumnName")?.Value ?? "AR";
                _generalConfig.EvidenceFontName = generalElement.Element("EvidenceFontName")?.Value ?? "MS PGothic";
                _generalConfig.BackButtonFontName = generalElement.Element("BackButtonFontName")?.Value ?? "Calibri";

                // Parse more numeric values
                if (int.TryParse(generalElement.Element("PrintAreaLastRowIdx")?.Value, out int lastRow))
                    _generalConfig.PrintAreaLastRowIdx = lastRow;

                if (double.TryParse(generalElement.Element("ColumnWidth")?.Value, out double colWidth))
                    _generalConfig.ColumnWidth = colWidth;

                if (double.TryParse(generalElement.Element("RowHeight")?.Value, out double rowHeight))
                    _generalConfig.RowHeight = rowHeight;

                if (int.TryParse(generalElement.Element("FontSize")?.Value, out int fontSize))
                    _generalConfig.FontSize = fontSize;

                // Parse page setup settings
                _generalConfig.PageOrientation = generalElement.Element("PageOrientation")?.Value ?? "Landscape";
                _generalConfig.PaperSize = generalElement.Element("PaperSize")?.Value ?? "A4";

                if (int.TryParse(generalElement.Element("Zoom")?.Value, out int zoom))
                    _generalConfig.Zoom = zoom;

                if (bool.TryParse(generalElement.Element("FitToPagesWide")?.Value, out bool fitWide))
                    _generalConfig.FitToPagesWide = fitWide;

                if (bool.TryParse(generalElement.Element("FitToPagesTall")?.Value, out bool fitTall))
                    _generalConfig.FitToPagesTall = fitTall;

                // Parse margin settings
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

            Logger.Info($"Loaded configuration: {_sheets.Count} sheets");
        }

        /// <summary>
        /// Tạo file cấu hình mặc định nếu chưa tồn tại
        /// Tạo XML structure với tất cả default values
        ///
        /// </summary>
        private static void CreateDefaultConfig()
        {
            _sheets = new List<SheetConfig>
      {
        new SheetConfig
        {
          Name = "共通",
          Prefix = "共通_",
          ReferenceColumnHeader = "参考 No.",
          NumberFormat = "D2",
          Description = "Sheet chung với prefix 共通_",
          IsHorizontal = "False"
        },
        new SheetConfig
        {
          Name = "テスト項目",
          Prefix = "エビデンス_",
          ReferenceColumnHeader = "参考 No.",
          NumberFormat = "D2",
          Description = "Sheet test items với prefix エビデンス_",
          IsHorizontal = "False"
        },
        new SheetConfig
        {
          Name = "設計書",
          Prefix = "設計_",
          ReferenceColumnHeader = "No.",
          NumberFormat = "D3",
          Description = "Sheet thiết kế với prefix 設計_",
          IsHorizontal = "False"
        },
        new SheetConfig
        {
          Name = "検証結果",
          Prefix = "検証_",
          ReferenceColumnHeader = "検証 No.",
          NumberFormat = "D2",
          Description = "Sheet kết quả kiểm chứng với prefix 検証_",
          IsHorizontal = "False"
        }
      };

            _generalConfig = new GeneralConfig();
            _loggingConfig = new LoggingConfig();

            // Tạo file XML mặc định với đầy đủ settings
            try
            {
                var doc = new XDocument(
                  new XElement("SheetConfiguration",
                    new XElement("Sheets",
                      from sheet in _sheets
                      select new XElement("Sheet",
                        new XElement("Name", sheet.Name),
                        new XElement("Prefix", sheet.Prefix),
                        new XElement("NumberFormat", sheet.NumberFormat),
                        new XElement("Description", sheet.Description),
                        new XElement("IsHorizontal", sheet.IsHorizontal)
                      )
                    ),
                    new XElement("ExtendedSheets"),
                    new XElement("GeneralSettings",
                      new XElement("HeaderRowIndex", _generalConfig.HeaderRowIndex),
                      new XElement("AutoFillCell", _generalConfig.AutoFillCell),
                      new XElement("EnableDebugLog", _generalConfig.EnableDebugLog),
                      new XElement("StartingNumber", _generalConfig.StartingNumber),
                      new XElement("PageBreakColumnName", _generalConfig.PageBreakColumnName),
                      new XElement("EvidenceFontName", _generalConfig.EvidenceFontName),
                      new XElement("BackButtonFontName", _generalConfig.BackButtonFontName),
                      new XElement("PrintAreaLastRowIdx", _generalConfig.PrintAreaLastRowIdx),
                      new XElement("ColumnWidth", _generalConfig.ColumnWidth),
                      new XElement("RowHeight", _generalConfig.RowHeight),
                      new XElement("FontSize", _generalConfig.FontSize),
                      new XElement("PageOrientation", _generalConfig.PageOrientation),
                      new XElement("PaperSize", _generalConfig.PaperSize),
                      new XElement("Zoom", _generalConfig.Zoom),
                      new XElement("FitToPagesWide", _generalConfig.FitToPagesWide),
                      new XElement("FitToPagesTall", _generalConfig.FitToPagesTall),
                      new XElement("LeftMargin", _generalConfig.LeftMargin),
                      new XElement("RightMargin", _generalConfig.RightMargin),
                      new XElement("TopMargin", _generalConfig.TopMargin),
                      new XElement("BottomMargin", _generalConfig.BottomMargin),
                      new XElement("HeaderMargin", _generalConfig.HeaderMargin),
                      new XElement("FooterMargin", _generalConfig.FooterMargin),
                      new XElement("CenterHorizontally", _generalConfig.CenterHorizontally),
                      new XElement("WindowZoom", _generalConfig.WindowZoom),
                      new XElement("ViewMode", _generalConfig.ViewMode)
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

                string configDir = GetAddInDirectory();
                Directory.CreateDirectory(configDir);
                string configPath = Path.Combine(configDir, ConfigFileName);
                doc.Save(configPath);

                Logger.Info($"Created default config file at {configPath}");
            }
            catch (Exception ex)
            {
                // Nếu không thể tạo file, chỉ log lỗi nhưng vẫn tiếp tục với config trong memory
                Logger.Warning($"Could not create default config file: {ex.Message}");
            }
        }

        #endregion

        #region Public Accessors

        /// <summary>
        /// Lấy cấu hình cho sheet theo tên
        /// Tìm kiếm trong special sheets trước, sau đó extended sheets
        ///
        /// </summary>
        /// <param name="sheetName">Tên sheet cần tìm config</param>
        /// <returns>SheetConfig hoặc null nếu không tìm thấy</returns>
        public static SheetConfig GetSheetConfig(string sheetName)
        {
            if (_sheets == null)
                LoadConfiguration();

            // Tìm trong special sheets trước
            var config = _sheets?.FirstOrDefault(s => s.Name == sheetName);
            if (config != null)
                return config;

            // Tìm trong extended sheets
            return null;
        }

        /// <summary>
        /// Lấy cấu hình chung
        /// Lazy load configuration nếu chưa được load
        ///
        /// </summary>
        /// <returns>GeneralConfig instance</returns>
        public static GeneralConfig GetGeneralConfig()
        {
            if (_generalConfig == null)
                LoadConfiguration();

            return _generalConfig ?? new GeneralConfig();
        }

        /// <summary>
        /// Lấy tất cả sheet configs
        /// Lazy load configuration nếu chưa được load
        ///
        /// </summary>
        /// <returns>Danh sách tất cả sheet configs</returns>
        public static List<SheetConfig> GetAllSheetConfigs()
        {
            if (_sheets == null)
                LoadConfiguration();

            Logger.Info($"GetAllSheetConfigs returning {_sheets?.Count ?? 0} sheets");
            return _sheets ?? new List<SheetConfig>();
        }

        /// <summary>
        /// Lấy tất cả special sheet configs (deprecated - use GetAllSheetConfigs instead)
        /// Lazy load configuration nếu chưa được load
        ///
        /// </summary>
        /// <returns>Danh sách special sheet configs</returns>
        [Obsolete("Use GetAllSheetConfigs() instead. SpecialSheets and ExtendedSheets have been merged.")]
        public static List<SheetConfig> GetSpecialSheets()
        {
            return GetAllSheetConfigs();
        }

        /// <summary>
        /// Lấy tất cả extended sheet configs (deprecated - use GetAllSheetConfigs instead)
        /// Lazy load configuration nếu chưa được load
        ///
        /// </summary>
        /// <returns>Danh sách extended sheet configs</returns>
        [Obsolete("Use GetAllSheetConfigs() instead. SpecialSheets and ExtendedSheets have been merged.")]
        public static List<SheetConfig> GetExtendedSheets()
        {
            return GetAllSheetConfigs();
        }

        /// <summary>
        /// Reload cấu hình từ file
        /// Force reload tất cả cached configurations
        ///
        /// </summary>
        public static void ReloadConfiguration()
        {
            _sheets = null;
            _generalConfig = null;
            _loggingConfig = null;
            LoadConfiguration();
        }

        /// <summary>
        /// Kiểm tra xem có sheet config nào cho tên sheet đã cho không
        ///
        /// </summary>
        /// <param name="sheetName">Tên sheet cần kiểm tra</param>
        /// <returns>true nếu có config, false nếu không</returns>
        public static bool HasSheetConfig(string sheetName)
        {
            return GetSheetConfig(sheetName) != null;
        }

        /// <summary>
        /// Lưu cấu hình vào file XML
        /// Serialize tất cả config objects thành XML và ghi ra file
        ///
        /// </summary>
        public static void SaveConfiguration()
        {
            try
            {
                var doc = new XDocument(
                  new XElement("SheetConfiguration",
                    new XElement("Sheets",
                      from sheet in _sheets ?? new List<SheetConfig>()
                      select new XElement("Sheet",
                        new XElement("Name", sheet.Name),
                        new XElement("Prefix", sheet.Prefix),
                        new XElement("ReferenceColumnHeader", sheet.ReferenceColumnHeader ?? ""),
                        new XElement("NumberFormat", sheet.NumberFormat),
                        new XElement("Description", sheet.Description),
                        new XElement("IsHorizontal", sheet.IsHorizontal)
                      )
                    ),
                    new XElement("GeneralSettings",
                      new XElement("HeaderRowIndex", _generalConfig?.HeaderRowIndex ?? 2),
                      new XElement("AutoFillCell", _generalConfig?.AutoFillCell ?? true),
                      new XElement("EnableDebugLog", _generalConfig?.EnableDebugLog ?? true),
                      new XElement("StartingNumber", _generalConfig?.StartingNumber ?? 1),
                      new XElement("PageBreakColumnName", _generalConfig?.PageBreakColumnName ?? "AR"),
                      new XElement("EvidenceFontName", _generalConfig?.EvidenceFontName ?? "MS PGothic"),
                      new XElement("BackButtonFontName", _generalConfig?.BackButtonFontName ?? "Calibri"),
                      new XElement("PrintAreaLastRowIdx", _generalConfig?.PrintAreaLastRowIdx ?? 38),
                      new XElement("ColumnWidth", _generalConfig?.ColumnWidth ?? 2.38),
                      new XElement("RowHeight", _generalConfig?.RowHeight ?? 12.6),
                      new XElement("FontSize", _generalConfig?.FontSize ?? 11),
                      new XElement("PageOrientation", _generalConfig?.PageOrientation ?? "Landscape"),
                      new XElement("PaperSize", _generalConfig?.PaperSize ?? "A4"),
                      new XElement("Zoom", _generalConfig?.Zoom ?? 100),
                      new XElement("FitToPagesWide", _generalConfig?.FitToPagesWide ?? false),
                      new XElement("FitToPagesTall", _generalConfig?.FitToPagesTall ?? false),
                      new XElement("LeftMargin", _generalConfig?.LeftMargin ?? 0.75),
                      new XElement("RightMargin", _generalConfig?.RightMargin ?? 0.75),
                      new XElement("TopMargin", _generalConfig?.TopMargin ?? 1.0),
                      new XElement("BottomMargin", _generalConfig?.BottomMargin ?? 1.0),
                      new XElement("HeaderMargin", _generalConfig?.HeaderMargin ?? 0.5),
                      new XElement("FooterMargin", _generalConfig?.FooterMargin ?? 0.5),
                      new XElement("CenterHorizontally", _generalConfig?.CenterHorizontally ?? true),
                      new XElement("WindowZoom", _generalConfig?.WindowZoom ?? 100),
                      new XElement("ViewMode", _generalConfig?.ViewMode ?? "PageBreakPreview")
                    ),
                    new XElement("LoggingSettings",
                      new XElement("LogDirectory", _loggingConfig?.LogDirectory ?? @"C:\ExcelCustomAddin"),
                      new XElement("EnableFileLogging", _loggingConfig?.EnableFileLogging ?? true),
                      new XElement("EnableDebugOutput", _loggingConfig?.EnableDebugOutput ?? true),
                      new XElement("LogLevel", _loggingConfig?.LogLevel ?? "DEBUG"),
                      new XElement("LogFileName", _loggingConfig?.LogFileName ?? "ExcelAddin")
                    )
                  )
                );

                string configDir = GetAddInDirectory();
                Directory.CreateDirectory(configDir);
                string configPath = Path.Combine(configDir, ConfigFileName);
                doc.Save(configPath);

                Logger.Info($"Saved configuration to {configPath}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error saving configuration: {ex.Message}", ex);
                throw;
            }
        }

        #endregion

        /// <summary>
        /// Lấy cấu hình logging
        /// Lazy load configuration nếu chưa được load
        ///
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
    }
}