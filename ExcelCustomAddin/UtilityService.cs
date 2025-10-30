using Microsoft.Office.Interop.Excel;
using System;
using System.Text.RegularExpressions;

namespace ExcelCustomAddin
{
  /// <summary>
  /// Service chứa các method tiện ích dùng chung
  /// </summary>
  public static class UtilityService
  {
    // Invalid characters for named ranges (Excel rules)
    private static readonly char[] INVALID_NAME_CHARACTERS = { '', '.', '~', ' ', '-', '+', '=', '*', '/', '\\', '[', ']', '(', ')', '{', '}', '<', '>', '!', '@', '#', '$', '%', '^', '&', '|', ':', ';', '"', '\'', ',', '?' };

    /// <summary>
    /// Get column index from column name (A=1, B=2, etc.)
    /// </summary>
    public static int GetColumnIndex(string columnName)
    {
      if (string.IsNullOrEmpty(columnName))
        return 1;

      columnName = columnName.ToUpper();
      int columnIndex = 0;
      int length = columnName.Length;

      for (int i = 0; i < length; i++)
      {
        columnIndex += (columnName[length - 1 - i] - 'A' + 1) * (int)Math.Pow(26, i);
      }

      return columnIndex;
    }

    /// <summary>
    /// Generate auto sheet name based on column and sheet type
    /// </summary>
    public static string GenerateAutoSheetName(Worksheet activeSheet, int column, string currentSheetName)
    {
      try
      {
        // Lấy config cho sheet hiện tại
        var sheetConfig = SheetConfigManager.GetSheetConfig(currentSheetName);
        if (sheetConfig == null)
          return null;

        // Lấy prefix từ config
        string prefix = sheetConfig.Prefix;
        if (string.IsNullOrEmpty(prefix))
          return null;

        // Lấy starting number từ config
        int startingNumber = SheetConfigManager.GetGeneralConfig().StartingNumber;

        // Tính số thứ tự dựa trên column
        int sequenceNumber = column + startingNumber - 1; // Column A = 1, so A + startingNumber - 1

        // Format số thứ tự theo config
        string numberFormat = sheetConfig.NumberFormat ?? "D2";
        string formattedNumber = sequenceNumber.ToString(numberFormat);

        // Tạo tên sheet
        string sheetName = $"{prefix}{formattedNumber}";

        // Kiểm tra tên có hợp lệ không
        if (IsValidNamedRangeName(sheetName))
        {
          return sheetName;
        }

        return null;
      }
      catch (Exception ex)
      {
        Logger.Warning($"Error generating auto sheet name: {ex.Message}");
        return null;
      }
    }

    /// <summary>
    /// Check if a named range name is valid according to Excel rules
    /// </summary>
    public static bool IsValidNamedRangeName(string name)
    {
      if (string.IsNullOrEmpty(name))
        return false;

      // Name cannot be longer than 255 characters
      if (name.Length > 255)
        return false;

      // Name cannot start with a number or contain invalid characters
      if (char.IsDigit(name[0]) || name.IndexOfAny(INVALID_NAME_CHARACTERS) >= 0)
        return false;

      // Name cannot be a cell reference (like A1, B2, etc.)
      if (Regex.IsMatch(name, @"^[A-Z]+[0-9]+$"))
        return false;

      // Name cannot be R or C (reserved for R1C1 reference style)
      if (name.Equals("R", StringComparison.OrdinalIgnoreCase) ||
          name.Equals("C", StringComparison.OrdinalIgnoreCase))
        return false;

      return true;
    }

    /// <summary>
    /// Generate a valid named range name from a cell address
    /// </summary>
    public static string GenerateValidRangeName(Range cell, Worksheet worksheet)
    {
      string baseName = $"Range_{worksheet.Name}_{cell.Address[false, false].Replace("$", "").Replace(":", "_")}";
      string validName = baseName;

      // Ensure name is valid and unique
      int counter = 1;
      while (!IsValidNamedRangeName(validName) || NameExistsInWorkbook(validName, worksheet.Application.ActiveWorkbook))
      {
        validName = $"{baseName}_{counter}";
        counter++;
        if (counter > 1000) // Prevent infinite loop
          break;
      }

      return IsValidNamedRangeName(validName) ? validName : null;
    }

    /// <summary>
    /// Ensure a named range name is unique in the workbook
    /// </summary>
    public static string EnsureUniqueRangeName(string baseName, Workbook workbook)
    {
      if (string.IsNullOrEmpty(baseName))
        return GenerateValidRangeName(null, null); // This will need to be handled differently

      string uniqueName = baseName;
      int counter = 1;

      while (NameExistsInWorkbook(uniqueName, workbook))
      {
        uniqueName = $"{baseName}_{counter}";
        counter++;
        if (counter > 1000) // Prevent infinite loop
          break;
      }

      return uniqueName;
    }

    /// <summary>
    /// Check if a named range exists in the workbook
    /// </summary>
    public static bool NameExistsInWorkbook(string name, Workbook workbook)
    {
      if (workbook == null || string.IsNullOrEmpty(name))
        return false;

      try
      {
        foreach (Name namedRange in workbook.Names)
        {
          if (namedRange.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
          {
            return true;
          }
        }
      }
      catch (Exception ex)
      {
        Logger.Warning($"Error checking if name exists: {ex.Message}");
      }

      return false;
    }

    /// <summary>
    /// Try to get active objects (workbook and worksheet)
    /// </summary>
    public static bool TryGetActiveObjects(out Workbook workbook, out Worksheet worksheet)
    {
      workbook = null;
      worksheet = null;

      try
      {
        var app = Globals.ThisAddIn.Application;
        workbook = app.ActiveWorkbook;
        worksheet = app.ActiveSheet as Worksheet;

        return workbook != null && worksheet != null;
      }
      catch
      {
        return false;
      }
    }

    /// <summary>
    /// Create or get named range for a cell
    /// </summary>
    public static string GetOrCreateNamedRangeForCell(Range cell, Worksheet worksheet)
    {
      if (cell == null || worksheet == null)
        return null;

      try
      {
        string cellAddress = cell.Address[false, false];
        string proposedName = $"BackTo_{worksheet.Name}_{cellAddress.Replace("$", "").Replace(":", "_")}";

        // Ensure name is valid and unique
        string validName = EnsureUniqueRangeName(proposedName, worksheet.Application.ActiveWorkbook);

        if (!string.IsNullOrEmpty(validName))
        {
          // Create the named range
          worksheet.Application.ActiveWorkbook.Names.Add(validName, cell, true);
          Logger.Debug($"Created named range '{validName}' for cell {cellAddress}");
          return validName;
        }
      }
      catch (Exception ex)
      {
        Logger.Warning($"Error creating named range for cell {cell?.Address[false, false]}: {ex.Message}");
      }

      return null;
    }

    /// <summary>
    /// Create named range with validation
    /// </summary>
    public static bool CreateNamedRange(Workbook workbook, string name, string address)
    {
      if (workbook == null || string.IsNullOrEmpty(name) || string.IsNullOrEmpty(address))
        return false;

      try
      {
        // Validate name
        if (!IsValidNamedRangeName(name))
        {
          Logger.Error($"Invalid named range name: {name}");
          return false;
        }

        // Ensure unique name
        string uniqueName = EnsureUniqueRangeName(name, workbook);

        // Create named range
        workbook.Names.Add(uniqueName, address, true);
        Logger.Info($"Created named range '{uniqueName}' with address '{address}'");
        return true;
      }
      catch (Exception ex)
      {
        Logger.Error($"Error creating named range '{name}': {ex.Message}", ex);
        return false;
      }
    }

    /// <summary>
    /// Create named range internal implementation
    /// </summary>
    public static bool CreateNamedRangeInternal(Workbook workbook, string name, Range range)
    {
      if (workbook == null || range == null || string.IsNullOrEmpty(name))
        return false;

      try
      {
        // Validate name
        if (!IsValidNamedRangeName(name))
        {
          Logger.Error($"Invalid named range name: {name}");
          return false;
        }

        // Ensure unique name
        string uniqueName = EnsureUniqueRangeName(name, workbook);

        // Create named range
        workbook.Names.Add(uniqueName, range, true);
        Logger.Info($"Created named range '{uniqueName}' for range {range.Address[false, false]}");
        return true;
      }
      catch (Exception ex)
      {
        Logger.Error($"Error creating named range '{name}': {ex.Message}", ex);
        return false;
      }
    }
  }
}