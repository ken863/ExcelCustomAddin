using System;
using System.IO;

namespace ExcelCustomAddin
{
  /// <summary>
  /// Simple logging utility for the Excel Custom Add-in
  /// </summary>
  public static class Logger
  {
    private static readonly object _lock = new object();
    private static string _logFilePath;
    private static bool _isDebugEnabled = true;
    private static string _customLogDirectory = null;

    static Logger()
    {
      InitializeLogPath();
    }

    /// <summary>
    /// Initialize the log file path
    /// </summary>
    private static void InitializeLogPath()
    {
      InitializeLogPath("ExcelAddin");
    }

    /// <summary>
    /// Set custom log directory path
    /// </summary>
    /// <param name="directoryPath">The directory path for log files</param>
    public static void SetLogDirectory(string directoryPath)
    {
      lock (_lock)
      {
        _customLogDirectory = directoryPath;
        InitializeLogPath();
      }
    }

    /// <summary>
    /// Set custom log file path (full path including filename)
    /// </summary>
    /// <param name="filePath">The full path for the log file</param>
    public static void SetLogFilePath(string filePath)
    {
      lock (_lock)
      {
        try
        {
          // Tạo thư mục nếu chưa tồn tại
          string directory = Path.GetDirectoryName(filePath);
          if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
          {
            Directory.CreateDirectory(directory);
          }

          _logFilePath = filePath;
        }
        catch (Exception ex)
        {
          // Nếu có lỗi, fallback về đường dẫn mặc định
          System.Diagnostics.Debug.WriteLine($"Error setting log file path: {ex.Message}");
          InitializeLogPath();
        }
      }
    }

    /// <summary>
    /// Configure logger from XML configuration
    /// </summary>
    /// <param name="logDirectory">Log directory from config</param>
    /// <param name="enableDebug">Enable debug logging from config</param>
    /// <param name="logFileName">Log file name prefix from config</param>
    public static void Configure(string logDirectory = null, bool enableDebug = true, string logFileName = "ExcelAddin")
    {
      lock (_lock)
      {
        _isDebugEnabled = enableDebug;

        if (!string.IsNullOrEmpty(logDirectory))
        {
          _customLogDirectory = logDirectory;
        }

        // Reinitialize with custom filename
        InitializeLogPath(logFileName);
      }
    }

    /// <summary>
    /// Initialize the log file path with custom filename
    /// </summary>
    /// <param name="fileName">Custom filename prefix</param>
    private static void InitializeLogPath(string fileName = "ExcelAddin")
    {
      string logDirectory;

      if (!string.IsNullOrEmpty(_customLogDirectory))
      {
        logDirectory = _customLogDirectory;
      }
      else
      {
        // Sử dụng thư mục mặc định C:\ExcelCustomAddin
        logDirectory = @"C:\ExcelCustomAddin";
      }

      // Tạo thư mục nếu chưa tồn tại
      if (!Directory.Exists(logDirectory))
      {
        try
        {
          Directory.CreateDirectory(logDirectory);
          System.Diagnostics.Debug.WriteLine($"Created log directory: {logDirectory}");
        }
        catch (Exception ex)
        {
          // Nếu không thể tạo thư mục C:\ExcelCustomAddin, fallback về Temp
          System.Diagnostics.Debug.WriteLine($"Failed to create log directory {logDirectory}: {ex.Message}");
          string tempPath = Path.GetTempPath();
          logDirectory = Path.Combine(tempPath, "ExcelCustomAddin");
          if (!Directory.Exists(logDirectory))
          {
            Directory.CreateDirectory(logDirectory);
          }
          System.Diagnostics.Debug.WriteLine($"Using fallback log directory: {logDirectory}");
        }
      }

      _logFilePath = Path.Combine(logDirectory, $"{fileName}_{DateTime.Now:yyyyMMdd}.log");
    }

    /// <summary>
    /// Enable or disable debug logging
    /// </summary>
    /// <param name="enabled">True to enable debug logging</param>
    public static void SetDebugEnabled(bool enabled)
    {
      _isDebugEnabled = enabled;
    }

    /// <summary>
    /// Write a debug message to the log
    /// </summary>
    /// <param name="message">The message to log</param>
    public static void Debug(string message)
    {
      if (_isDebugEnabled)
      {
        WriteLog("DEBUG", message);
      }
    }

    /// <summary>
    /// Write an info message to the log
    /// </summary>
    /// <param name="message">The message to log</param>
    public static void Info(string message)
    {
      WriteLog("INFO", message);
    }

    /// <summary>
    /// Write a warning message to the log
    /// </summary>
    /// <param name="message">The message to log</param>
    public static void Warning(string message)
    {
      WriteLog("WARNING", message);
    }

    /// <summary>
    /// Write an error message to the log
    /// </summary>
    /// <param name="message">The message to log</param>
    public static void Error(string message)
    {
      WriteLog("ERROR", message);
    }

    /// <summary>
    /// Write an error message with exception details to the log
    /// </summary>
    /// <param name="message">The message to log</param>
    /// <param name="ex">The exception to log</param>
    public static void Error(string message, Exception ex)
    {
      WriteLog("ERROR", $"{message}: {ex.Message}\nStack Trace: {ex.StackTrace}");
    }

    /// <summary>
    /// Write a formatted message to the log file
    /// </summary>
    /// <param name="level">Log level (DEBUG, INFO, WARNING, ERROR)</param>
    /// <param name="message">The message to log</param>
    private static void WriteLog(string level, string message)
    {
      try
      {
        lock (_lock)
        {
          string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
          string logEntry = $"[{timestamp}] [{level}] {message}";

          // Write to file
          File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);

          // Also write to Debug output for development
          System.Diagnostics.Debug.WriteLine(logEntry);
        }
      }
      catch (Exception ex)
      {
        // Fallback to Debug output if file logging fails
        System.Diagnostics.Debug.WriteLine($"Logger Error: {ex.Message}");
        System.Diagnostics.Debug.WriteLine($"Original message: [{level}] {message}");
      }
    }

    /// <summary>
    /// Get the current log file path
    /// </summary>
    /// <returns>The path to the current log file</returns>
    public static string GetLogFilePath()
    {
      return _logFilePath;
    }

    /// <summary>
    /// Clear the current log file
    /// </summary>
    public static void ClearLog()
    {
      try
      {
        lock (_lock)
        {
          if (File.Exists(_logFilePath))
          {
            File.WriteAllText(_logFilePath, string.Empty);
          }
        }
      }
      catch (Exception ex)
      {
        System.Diagnostics.Debug.WriteLine($"Error clearing log file: {ex.Message}");
      }
    }
  }
}
