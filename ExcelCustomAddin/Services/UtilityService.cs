using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
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
    /// Create or get named range for a cell using cell value as name
    /// Handles special cell types and provides robust fallback logic
    /// </summary>
    public static string GetOrCreateNamedRangeForCell(Range cell, Worksheet worksheet)
    {
      if (cell == null || worksheet == null)
        return null;

      try
      {
        string cellAddress = cell.Address[false, false];

        // Determine cell type and extract appropriate value
        string cellValue = ExtractCellValueForNaming(cell);

        // Sanitize cell value để tạo tên hợp lệ
        string proposedName = SanitizeForNamedRange(cellValue);

        // Multiple fallback strategies if sanitization fails
        if (string.IsNullOrEmpty(proposedName))
        {
          // Fallback 1: Try to use cell address
          proposedName = SanitizeForNamedRange(cellAddress.Replace("$", "").Replace(":", "_"));
        }

        if (string.IsNullOrEmpty(proposedName))
        {
          // Fallback 2: Use generic cell reference
          proposedName = $"Cell_{cellAddress.Replace("$", "").Replace(":", "_")}";
        }

        // Ensure name is valid and unique
        string validName = EnsureUniqueRangeName(proposedName, worksheet.Application.ActiveWorkbook);

        if (!string.IsNullOrEmpty(validName))
        {
          // Create the named range
          worksheet.Application.ActiveWorkbook.Names.Add(validName, cell, true);
          Logger.Debug($"Created named range '{validName}' for cell {cellAddress} with value '{cellValue}'");
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
    /// Extract appropriate value from cell for naming purposes
    /// Handles different cell types intelligently
    /// </summary>
    public static string ExtractCellValueForNaming(Range cell)
    {
      if (cell == null)
        return string.Empty;

      try
      {
        // Check if cell has a formula
        if (!string.IsNullOrEmpty(cell.Formula?.ToString()) && cell.Formula.ToString().StartsWith("="))
        {
          // For formula cells, try to get displayed text or use formula type
          string displayedText = cell.Text?.ToString()?.Trim();
          if (!string.IsNullOrEmpty(displayedText) && displayedText != cell.Formula.ToString())
          {
            return displayedText;
          }
          // If formula result is same as formula, classify by formula type
          return ClassifyFormulaType(cell.Formula.ToString());
        }

        // Try to get cell value
        object cellValue = cell.Value2;
        if (cellValue != null)
        {
          string stringValue = cellValue.ToString().Trim();

          // Handle different data types
          if (double.TryParse(stringValue, out double numValue))
          {
            return numValue.ToString(); // Keep as number for SanitizeForNamedRange to handle
          }

          if (bool.TryParse(stringValue, out bool boolValue))
          {
            return boolValue.ToString().ToUpper(); // TRUE or FALSE
          }

          if (DateTime.TryParse(stringValue, out DateTime dateValue))
          {
            return dateValue.ToString("yyyy-MM-dd"); // Standardized date format
          }

          return stringValue;
        }

        // If no value, try to get displayed text
        string textValue = cell.Text?.ToString()?.Trim();
        if (!string.IsNullOrEmpty(textValue))
        {
          return textValue;
        }

        // Last resort: empty string
        return string.Empty;
      }
      catch (Exception ex)
      {
        Logger.Warning($"Error extracting value from cell {cell?.Address[false, false]}: {ex.Message}");
        return string.Empty;
      }
    }

    /// <summary>
    /// Classify formula type for naming purposes
    /// </summary>
    private static string ClassifyFormulaType(string formula)
    {
      if (string.IsNullOrEmpty(formula))
        return "Formula";

      string upperFormula = formula.ToUpper();

      if (upperFormula.Contains("SUM(")) return "Sum_Formula";
      if (upperFormula.Contains("COUNT(")) return "Count_Formula";
      if (upperFormula.Contains("AVERAGE(")) return "Average_Formula";
      if (upperFormula.Contains("IF(")) return "Conditional_Formula";
      if (upperFormula.Contains("VLOOKUP(")) return "Lookup_Formula";
      if (upperFormula.Contains("INDEX(")) return "Index_Formula";
      if (upperFormula.Contains("MATCH(")) return "Match_Formula";

      return "Formula";
    }
    public static string SanitizeForNamedRange(string input)
    {
      if (string.IsNullOrEmpty(input))
        return string.Empty;

      string sanitized = input.Trim();

      // Handle special cell value types
      if (sanitized.StartsWith("="))
      {
        // Formula - extract meaningful part or use generic name
        return "Formula_Range";
      }

      // Handle boolean values
      if (sanitized.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
        return "Bool_True";
      if (sanitized.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
        return "Bool_False";

      // Handle pure numeric values
      if (double.TryParse(sanitized, out double numericValue))
      {
        return $"Num_{numericValue}";
      }

      // Handle date/time values (if they come as strings)
      if (DateTime.TryParse(sanitized, out DateTime dateValue))
      {
        return $"Date_{dateValue:yyyy_MM_dd}";
      }

      // Check for Excel reserved keywords
      string[] excelKeywords = { "R", "C", "TRUE", "FALSE", "AND", "OR", "NOT", "IF", "SUM", "COUNT", "AVERAGE", "MIN", "MAX" };
      if (Array.Exists(excelKeywords, keyword => keyword.Equals(sanitized, StringComparison.OrdinalIgnoreCase)))
      {
        return $"{sanitized}_Range";
      }

      // Replace invalid characters with underscores
      sanitized = Regex.Replace(sanitized, @"[^a-zA-Z0-9_]", "_");

      // Remove multiple consecutive underscores
      sanitized = Regex.Replace(sanitized, @"_+", "_");

      // Remove leading/trailing underscores and numbers
      sanitized = sanitized.Trim('_');

      // If empty after sanitization, return empty (will use fallback)
      if (string.IsNullOrEmpty(sanitized))
        return string.Empty;

      // Ensure it starts with a letter (not a number or underscore)
      if (char.IsDigit(sanitized[0]) || sanitized[0] == '_')
      {
        sanitized = $"R_{sanitized}";
      }

      // Handle very short names (less than 2 characters)
      if (sanitized.Length < 2)
      {
        sanitized = $"Short_{sanitized}";
      }

      // Limit length to 255 characters and ensure it doesn't end with underscore
      if (sanitized.Length > 255)
      {
        sanitized = sanitized.Substring(0, 255).TrimEnd('_');
      }

      // Final check - ensure it's not empty and valid
      return IsValidNamedRangeName(sanitized) ? sanitized : string.Empty;
    }
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