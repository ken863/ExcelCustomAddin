using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelCustomAddin
{
    /// <summary>
    /// UtilityService - Service chứa các method tiện ích dùng chung cho Excel Add-in
    ///
    /// Chức năng chính:
    /// - Xử lý tên cột và chỉ số cột Excel
    /// - Tạo tên sheet tự động dựa trên cấu hình
    /// - Validation và xử lý named range
    /// - Sanitize tên cho named range từ giá trị cell
    /// - Quản lý named range trong workbook
    ///
    /// Tác giả: lam.pt
    /// Ngày tạo: 2025
    /// </summary>
    public static class UtilityService
    {
        #region Constants

        /// <summary>
        /// Các ký tự không hợp lệ trong named range của Excel
        /// Theo quy tắc của Excel: không được chứa các ký tự đặc biệt này
        /// </summary>
        private static readonly char[] INVALID_NAME_CHARACTERS = {
      '', '.', '~', ' ', '-', '+', '=', '*', '/', '\\',
      '[', ']', '(', ')', '{', '}', '<', '>', '!',
      '@', '#', '$', '%', '^', '&', '|', ':', ';',
      '"', '\'', ',', '?'
    };

        #endregion

        #region Excel Column Operations

        /// <summary>
        /// Chuyển đổi tên cột Excel thành chỉ số cột (A=1, B=2, AA=27, etc.)
        ///
        /// Ví dụ:
        /// - GetColumnIndex("A") = 1
        /// - GetColumnIndex("Z") = 26
        /// - GetColumnIndex("AA") = 27
        /// - GetColumnIndex("AZ") = 52
        ///
        /// </summary>
        /// <param name="columnName">Tên cột Excel (A, B, AA, AZ, etc.)</param>
        /// <returns>Chỉ số cột (bắt đầu từ 1)</returns>
        public static int GetColumnIndex(string columnName)
        {
            if (string.IsNullOrEmpty(columnName))
                return 1;

            columnName = columnName.ToUpper();
            int columnIndex = 0;
            int length = columnName.Length;

            // Chuyển đổi từ hệ cơ số 26 (A-Z) sang hệ thập phân
            for (int i = 0; i < length; i++)
            {
                columnIndex += (columnName[length - 1 - i] - 'A' + 1) * (int)Math.Pow(26, i);
            }

            return columnIndex;
        }

        #endregion

        #region Sheet Name Generation

        /// <summary>
        /// Tạo tên sheet tự động dựa trên vị trí cột và cấu hình sheet
        ///
        /// Quy tắc tạo tên:
        /// 1. Lấy prefix từ cấu hình sheet
        /// 2. Tính số thứ tự dựa trên giá trị lớn nhất trong cột (loại bỏ prefix)
        /// 3. Format số thứ tự theo cấu hình
        /// 4. Kết hợp prefix + số thứ tự
        ///
        /// Ví dụ: Prefix="INV", cột A có ["INV001", "INV002"]
        /// → Sequence number = 3, Result: "INV003"
        /// </summary>
        /// <param name="activeSheet">Sheet hiện tại</param>
        /// <param name="column">Vị trí cột (1-based)</param>
        /// <param name="currentSheetName">Tên sheet hiện tại để lấy cấu hình</param>
        /// <returns>Tên sheet đã tạo hoặc null nếu không thể tạo</returns>
        public static string GenerateAutoSheetName(Worksheet activeSheet, int column, string currentSheetName)
        {
            try
            {
                // Lấy cấu hình cho sheet hiện tại từ SheetConfigManager
                var sheetConfig = SheetConfigManager.GetSheetConfig(currentSheetName);
                if (sheetConfig == null)
                {
                    Logger.Warning($"Không tìm thấy cấu hình cho sheet: {currentSheetName}");
                    return null;
                }

                // Lấy prefix từ cấu hình (bắt buộc)
                string prefix = sheetConfig.Prefix;
                if (string.IsNullOrEmpty(prefix))
                {
                    Logger.Warning($"Prefix không được cấu hình cho sheet: {currentSheetName}");
                    return null;
                }

                // Tính số thứ tự dựa trên giá trị lớn nhất trong cột (loại bỏ prefix)
                // Thay vì dùng column index, scan cột để tìm sequence number cao nhất
                int sequenceNumber = GetNextSequenceNumber(activeSheet, column, prefix);

                // Format số thứ tự theo cấu hình (VD: "D2" = 2 chữ số, "D3" = 3 chữ số)
                string numberFormat = sheetConfig.NumberFormat ?? "D2";
                string formattedNumber = sequenceNumber.ToString(numberFormat);

                // Tạo tên sheet cuối cùng
                string sheetName = $"{prefix}{formattedNumber}";

                // Kiểm tra tính hợp lệ của tên
                if (IsValidNamedRangeName(sheetName))
                {
                    Logger.Debug($"Generated sheet name: {sheetName} (column: {column}, prefix: {prefix}, sequence: {sequenceNumber})");
                    return sheetName;
                }
                else
                {
                    Logger.Warning($"Tên sheet được tạo không hợp lệ: {sheetName}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Lỗi khi tạo tên sheet tự động: {ex.Message}");
                return null;
            }
        }
        /// Tính sequence number tiếp theo dựa trên giá trị lớn nhất trong cột
        /// Scan cột để tìm tất cả giá trị có prefix, trích xuất số và lấy max + 1
        ///
        /// Ví dụ: Prefix="INV", cột có giá trị ["INV001", "INV002", "ABC123"]
        /// → Tìm thấy [1, 2], max = 2, trả về 3
        ///
        /// </summary>
        /// <param name="activeSheet">Sheet chứa cột cần scan</param>
        /// <param name="column">Vị trí cột (1-based)</param>
        /// <param name="prefix">Prefix để lọc giá trị</param>
        /// <returns>Sequence number tiếp theo</returns>
        private static int GetNextSequenceNumber(Worksheet activeSheet, int column, string prefix)
        {
            try
            {
                int maxSequence = 0;
                int startingNumber = SheetConfigManager.GetGeneralConfig()?.StartingNumber ?? 1;

                // Tìm cell cuối cùng có giá trị trong cột
                Range lastCell = null;
                try
                {
                    // Sử dụng UsedRange để tìm vùng đã sử dụng, sau đó lấy cell cuối cùng của cột
                    var usedRange = activeSheet.UsedRange;
                    var lastRow = usedRange.Row + usedRange.Rows.Count - 1;
                    var lastCol = usedRange.Column + usedRange.Columns.Count - 1;

                    // Nếu cột cần tìm nằm ngoài UsedRange, sử dụng End(xlUp) từ cell cuối cùng của sheet
                    if (column > lastCol)
                    {
                        lastCell = activeSheet.Cells[activeSheet.Rows.Count, column].End(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                    }
                    else
                    {
                        // Tìm cell cuối cùng có giá trị trong cột bằng End(xlUp) từ cell cuối của UsedRange
                        lastCell = activeSheet.Cells[lastRow, column].End(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                    }

                    // Nếu không tìm thấy cell nào có giá trị, lastCell sẽ là row 1
                    if (lastCell.Row == 1 && string.IsNullOrEmpty(lastCell.Value2?.ToString()))
                    {
                        Logger.Debug($"Column {column} appears to be empty, using starting number: {startingNumber}");
                        return startingNumber;
                    }
                }
                catch
                {
                    // Fallback: sử dụng End(xlUp) từ cell cuối cùng của sheet
                    lastCell = activeSheet.Cells[activeSheet.Rows.Count, column].End(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
                    if (lastCell.Row == 1 && string.IsNullOrEmpty(lastCell.Value2?.ToString()))
                    {
                        Logger.Debug($"Column {column} appears to be empty, using starting number: {startingNumber}");
                        return startingNumber;
                    }
                }

                // Lấy range từ row 1 đến row của lastCell
                Range columnRange = activeSheet.Range[activeSheet.Cells[1, column], lastCell];

                // Duyệt qua từng cell trong range đã xác định
                foreach (Range cell in columnRange.Cells)
                {
                    string cellValue = cell.Value2?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(cellValue))
                        continue;

                    // Kiểm tra xem cell có bắt đầu bằng prefix không
                    if (cellValue.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    {
                        // Trích xuất phần số sau prefix
                        string numberPart = cellValue.Substring(prefix.Length);
                        if (int.TryParse(numberPart, out int sequenceValue))
                        {
                            maxSequence = Math.Max(maxSequence, sequenceValue);
                        }
                    }
                }

                // Trả về max + 1, hoặc starting number nếu không tìm thấy giá trị nào
                int nextSequence = maxSequence > 0 ? maxSequence + 1 : startingNumber;

                Logger.Debug($"Next sequence number for column {column}, prefix '{prefix}': {nextSequence} (max found: {maxSequence})");
                return nextSequence;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error calculating next sequence number: {ex.Message}");
                // Fallback về starting number
                return SheetConfigManager.GetGeneralConfig()?.StartingNumber ?? 1;
            }
        }

        #endregion

        #region Named Range Validation

        /// <summary>
        /// Kiểm tra xem một tên có hợp lệ để làm named range trong Excel không
        ///
        /// Quy tắc validation của Excel:
        /// - Tối đa 255 ký tự
        /// - Bắt đầu bằng chữ cái hoặc underscore (không được bắt đầu bằng số)
        /// - Không chứa ký tự đặc biệt
        /// - Không trùng với địa chỉ cell (A1, B2, etc.)
        /// - Không phải là R hoặc C (dành cho R1C1 reference style)
        ///
        /// </summary>
        /// <param name="name">Tên cần kiểm tra</param>
        /// <returns>true nếu tên hợp lệ, false nếu không</returns>
        public static bool IsValidNamedRangeName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return false;

            // Rule 1: Tối đa 255 ký tự
            if (name.Length > 255)
                return false;

            // Rule 2: Không bắt đầu bằng số
            if (char.IsDigit(name[0]))
                return false;

            // Rule 3: Không chứa ký tự không hợp lệ
            if (name.IndexOfAny(INVALID_NAME_CHARACTERS) >= 0)
                return false;

            // Rule 4: Không phải địa chỉ cell (A1, B2, AA10, etc.)
            if (Regex.IsMatch(name, @"^[A-Z]+[0-9]+$"))
                return false;

            // Rule 5: Không phải R hoặc C (reserved for R1C1 style)
            if (name.Equals("R", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("C", StringComparison.OrdinalIgnoreCase))
                return false;

            return true;
        }

        #endregion

        #region Named Range Sanitization

        /// <summary>
        /// Làm sạch một chuỗi để có thể sử dụng làm tên named range
        /// Xử lý các trường hợp đặc biệt như công thức, số, ngày tháng, boolean, từ khóa Excel
        ///
        /// Quy trình xử lý:
        /// 1. Xử lý các loại giá trị đặc biệt (formula, boolean, number, date)
        /// 2. Thay thế ký tự không hợp lệ bằng underscore
        /// 3. Loại bỏ underscore liên tiếp
        /// 4. Đảm bảo bắt đầu bằng chữ cái
        /// 5. Giới hạn độ dài
        /// 6. Kiểm tra tính hợp lệ cuối cùng
        ///
        /// </summary>
        /// <param name="input">Chuỗi đầu vào cần làm sạch</param>
        /// <returns>Chuỗi đã được làm sạch và hợp lệ, hoặc empty string nếu không thể xử lý</returns>
        public static string SanitizeForNamedRange(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;

            string sanitized = input.Trim();

            // Xử lý các loại giá trị đặc biệt

            // 1. Công thức (bắt đầu bằng =)
            if (sanitized.StartsWith("="))
            {
                return "Formula_Range";
            }

            // 2. Giá trị boolean
            if (sanitized.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
                return "Bool_True";
            if (sanitized.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                return "Bool_False";

            // 3. Giá trị số thuần túy
            if (double.TryParse(sanitized, out double numericValue))
            {
                return $"Num_{numericValue}";
            }

            // 4. Giá trị ngày tháng
            if (DateTime.TryParse(sanitized, out DateTime dateValue))
            {
                return $"Date_{dateValue:yyyy_MM_dd}";
            }

            // 5. Từ khóa Excel reserved
            string[] excelKeywords = { "R", "C", "TRUE", "FALSE", "AND", "OR", "NOT", "IF", "SUM", "COUNT", "AVERAGE", "MIN", "MAX" };
            if (Array.Exists(excelKeywords, keyword => keyword.Equals(sanitized, StringComparison.OrdinalIgnoreCase)))
            {
                return $"{sanitized}_Range";
            }

            // 6. Thay thế ký tự không hợp lệ bằng underscore
            sanitized = Regex.Replace(sanitized, @"[^a-zA-Z0-9_]", "_");

            // 7. Loại bỏ nhiều underscore liên tiếp
            sanitized = Regex.Replace(sanitized, @"_+", "_");

            // 8. Loại bỏ underscore đầu và cuối
            sanitized = sanitized.Trim('_');

            // 9. Nếu rỗng sau khi làm sạch, trả về empty
            if (string.IsNullOrEmpty(sanitized))
                return string.Empty;

            // 10. Đảm bảo bắt đầu bằng chữ cái (không phải số hoặc underscore)
            if (char.IsDigit(sanitized[0]) || sanitized[0] == '_')
            {
                sanitized = $"R_{sanitized}";
            }

            // 11. Xử lý tên quá ngắn (< 2 ký tự)
            if (sanitized.Length < 2)
            {
                sanitized = $"Short_{sanitized}";
            }

            // 12. Giới hạn độ dài và loại bỏ underscore cuối
            if (sanitized.Length > 255)
            {
                sanitized = sanitized.Substring(0, 255).TrimEnd('_');
            }

            // 13. Kiểm tra tính hợp lệ cuối cùng
            return IsValidNamedRangeName(sanitized) ? sanitized : string.Empty;
        }

        /// <summary>
        /// Trích xuất giá trị phù hợp từ cell để đặt tên named range
        /// Xử lý thông minh các loại cell khác nhau (text, number, formula, date, boolean)
        ///
        /// Ưu tiên:
        /// 1. Giá trị hiển thị (đối với công thức)
        /// 2. Giá trị gốc của cell
        /// 3. Text hiển thị
        /// 4. Fallback về empty
        ///
        /// </summary>
        /// <param name="cell">Cell Excel cần trích xuất giá trị</param>
        /// <returns>Giá trị string phù hợp để đặt tên</returns>
        public static string ExtractCellValueForNaming(Range cell)
        {
            if (cell == null)
                return string.Empty;

            try
            {
                // Ưu tiên 1: Kiểm tra công thức
                if (!string.IsNullOrEmpty(cell.Formula?.ToString()) && cell.Formula.ToString().StartsWith("="))
                {
                    // Với cell có công thức, thử lấy text hiển thị
                    string displayedText = cell.Text?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(displayedText) && displayedText != cell.Formula.ToString())
                    {
                        return displayedText; // Ưu tiên kết quả hiển thị
                    }

                    // Nếu không có kết quả khác biệt, phân loại theo loại công thức
                    return ClassifyFormulaType(cell.Formula.ToString());
                }

                // Ưu tiên 2: Lấy giá trị gốc của cell
                object cellValue = cell.Value2;
                if (cellValue != null)
                {
                    string stringValue = cellValue.ToString().Trim();

                    // Xử lý các kiểu dữ liệu khác nhau
                    if (double.TryParse(stringValue, out double numValue))
                    {
                        return numValue.ToString(); // Giữ nguyên số để SanitizeForNamedRange xử lý
                    }

                    if (bool.TryParse(stringValue, out bool boolValue))
                    {
                        return boolValue.ToString().ToUpper(); // TRUE hoặc FALSE
                    }

                    if (DateTime.TryParse(stringValue, out DateTime dateValue))
                    {
                        return dateValue.ToString("yyyy-MM-dd"); // Format ngày chuẩn
                    }

                    return stringValue; // Text thông thường
                }

                // Ưu tiên 3: Lấy text hiển thị
                string textValue = cell.Text?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(textValue))
                {
                    return textValue;
                }

                // Cuối cùng: trả về empty
                return string.Empty;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Lỗi khi trích xuất giá trị từ cell {cell?.Address[false, false]}: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Phân loại loại công thức để tạo tên có ý nghĩa
        /// Dựa trên các hàm Excel phổ biến
        /// </summary>
        /// <param name="formula">Công thức Excel (VD: "=SUM(A1:A10)")</param>
        /// <returns>Tên phân loại công thức</returns>
        private static string ClassifyFormulaType(string formula)
        {
            if (string.IsNullOrEmpty(formula))
                return "Formula";

            string upperFormula = formula.ToUpper();

            // Các hàm toán học
            if (upperFormula.Contains("SUM(")) return "Sum_Formula";
            if (upperFormula.Contains("COUNT(")) return "Count_Formula";
            if (upperFormula.Contains("AVERAGE(")) return "Average_Formula";
            if (upperFormula.Contains("MIN(")) return "Min_Formula";
            if (upperFormula.Contains("MAX(")) return "Max_Formula";

            // Các hàm logic
            if (upperFormula.Contains("IF(")) return "Conditional_Formula";
            if (upperFormula.Contains("AND(")) return "Logic_AND_Formula";
            if (upperFormula.Contains("OR(")) return "Logic_OR_Formula";

            // Các hàm lookup
            if (upperFormula.Contains("VLOOKUP(")) return "Lookup_Formula";
            if (upperFormula.Contains("HLOOKUP(")) return "Lookup_Formula";
            if (upperFormula.Contains("INDEX(")) return "Index_Formula";
            if (upperFormula.Contains("MATCH(")) return "Match_Formula";

            // Các hàm text
            if (upperFormula.Contains("CONCATENATE(") || upperFormula.Contains("CONCAT(")) return "Text_Formula";
            if (upperFormula.Contains("LEFT(")) return "Text_Formula";
            if (upperFormula.Contains("RIGHT(")) return "Text_Formula";
            if (upperFormula.Contains("MID(")) return "Text_Formula";

            // Các hàm date/time
            if (upperFormula.Contains("TODAY(")) return "Date_Formula";
            if (upperFormula.Contains("NOW(")) return "DateTime_Formula";
            if (upperFormula.Contains("DATE(")) return "Date_Formula";

            // Fallback
            return "Formula";
        }

        #endregion

        #region Named Range Management

        /// <summary>
        /// Tạo hoặc lấy named range cho một cell, sử dụng giá trị cell làm tên
        /// Xử lý thông minh các loại cell đặc biệt và cung cấp fallback robust
        ///
        /// Quy trình:
        /// 1. Trích xuất giá trị từ cell
        /// 2. Sanitize để tạo tên hợp lệ
        /// 3. Đảm bảo tên unique trong workbook
        /// 4. Tạo named range
        ///
        /// </summary>
        /// <param name="cell">Cell Excel cần tạo named range</param>
        /// <param name="worksheet">Worksheet chứa cell</param>
        /// <returns>Tên named range đã tạo, hoặc null nếu thất bại</returns>
        public static string GetOrCreateNamedRangeForCell(Range cell, Worksheet worksheet)
        {
            if (cell == null || worksheet == null)
                return null;

            try
            {
                string cellAddress = cell.Address[false, false];

                // Bước 1: Trích xuất giá trị từ cell
                string cellValue = ExtractCellValueForNaming(cell);

                // Bước 2: Sanitize để tạo tên hợp lệ
                string proposedName = SanitizeForNamedRange(cellValue);

                // Bước 3: Fallback strategies nếu sanitization thất bại
                if (string.IsNullOrEmpty(proposedName))
                {
                    // Fallback 1: Thử dùng địa chỉ cell
                    proposedName = SanitizeForNamedRange(cellAddress.Replace("$", "").Replace(":", "_"));
                }

                if (string.IsNullOrEmpty(proposedName))
                {
                    // Fallback 2: Dùng tên generic với địa chỉ cell
                    proposedName = $"Cell_{cellAddress.Replace("$", "").Replace(":", "_")}";
                }

                // Bước 4: Đảm bảo tên unique trong workbook
                string validName = EnsureUniqueRangeName(proposedName, worksheet.Application.ActiveWorkbook);

                if (!string.IsNullOrEmpty(validName))
                {
                    // Bước 5: Tạo named range
                    worksheet.Application.ActiveWorkbook.Names.Add(validName, cell, true);
                    Logger.Debug($"Đã tạo named range '{validName}' cho cell {cellAddress} với giá trị '{cellValue}'");
                    return validName;
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Lỗi khi tạo named range cho cell {cell?.Address[false, false]}: {ex.Message}");
            }

            return null;
        }

        /// <summary>
        /// Đảm bảo tên named range là duy nhất trong workbook
        /// Nếu tên đã tồn tại, thêm hậu tố số (name_1, name_2, etc.)
        ///
        /// </summary>
        /// <param name="baseName">Tên cơ sở cần đảm bảo unique</param>
        /// <param name="workbook">Workbook cần kiểm tra</param>
        /// <returns>Tên unique hoặc null nếu không thể tạo</returns>
        public static string EnsureUniqueRangeName(string baseName, Workbook workbook)
        {
            if (string.IsNullOrEmpty(baseName))
                return GenerateValidRangeName(null, null); // Sẽ được xử lý khác

            string uniqueName = baseName;
            int counter = 1;

            // Thử thêm hậu tố số cho đến khi tìm được tên unique
            while (NameExistsInWorkbook(uniqueName, workbook))
            {
                uniqueName = $"{baseName}_{counter}";
                counter++;

                // Tránh vòng lặp vô hạn
                if (counter > 1000)
                    break;
            }

            return uniqueName;
        }

        /// <summary>
        /// Kiểm tra xem named range có tồn tại trong workbook không
        /// Duyệt qua tất cả names trong workbook để tìm tên trùng khớp
        ///
        /// </summary>
        /// <param name="name">Tên cần kiểm tra</param>
        /// <param name="workbook">Workbook cần kiểm tra</param>
        /// <returns>true nếu tên đã tồn tại, false nếu chưa</returns>
        public static bool NameExistsInWorkbook(string name, Workbook workbook)
        {
            if (workbook == null || string.IsNullOrEmpty(name))
                return false;

            try
            {
                // Duyệt qua tất cả named ranges trong workbook
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
                Logger.Warning($"Lỗi khi kiểm tra tên tồn tại: {ex.Message}");
            }

            return false;
        }

        /// <summary>
        /// Tạo named range với validation đầy đủ
        /// Kết hợp validation, uniqueness check và creation
        ///
        /// </summary>
        /// <param name="workbook">Workbook chứa named range</param>
        /// <param name="name">Tên named range</param>
        /// <param name="address">Địa chỉ range (VD: "A1" hoặc "A1:B10")</param>
        /// <returns>true nếu tạo thành công, false nếu thất bại</returns>
        public static bool CreateNamedRange(Workbook workbook, string name, string address)
        {
            if (workbook == null || string.IsNullOrEmpty(name) || string.IsNullOrEmpty(address))
                return false;

            try
            {
                // Validation tên
                if (!IsValidNamedRangeName(name))
                {
                    Logger.Error($"Tên named range không hợp lệ: {name}");
                    return false;
                }

                // Đảm bảo unique
                string uniqueName = EnsureUniqueRangeName(name, workbook);

                // Tạo named range
                workbook.Names.Add(uniqueName, address, true);
                Logger.Info($"Đã tạo named range '{uniqueName}' với địa chỉ '{address}'");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"Lỗi khi tạo named range '{name}': {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// Tạo named range từ Range object với validation đầy đủ
        /// Version internal với Range object thay vì string address
        ///
        /// </summary>
        /// <param name="workbook">Workbook chứa named range</param>
        /// <param name="name">Tên named range</param>
        /// <param name="range">Range object của Excel</param>
        /// <returns>true nếu tạo thành công, false nếu thất bại</returns>
        public static bool CreateNamedRangeInternal(Workbook workbook, string name, Range range)
        {
            if (workbook == null || range == null || string.IsNullOrEmpty(name))
                return false;

            try
            {
                // Validation tên
                if (!IsValidNamedRangeName(name))
                {
                    Logger.Error($"Tên named range không hợp lệ: {name}");
                    return false;
                }

                // Đảm bảo unique
                string uniqueName = EnsureUniqueRangeName(name, workbook);

                // Tạo named range từ Range object
                workbook.Names.Add(uniqueName, range, true);
                Logger.Info($"Đã tạo named range '{uniqueName}' cho range {range.Address[false, false]}");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"Lỗi khi tạo named range '{name}': {ex.Message}", ex);
                return false;
            }
        }

        /// <summary>
        /// Tạo tên named range hợp lệ từ địa chỉ cell (fallback method)
        /// Sử dụng khi không thể tạo tên từ giá trị cell
        ///
        /// </summary>
        /// <param name="cell">Cell để lấy địa chỉ</param>
        /// <param name="worksheet">Worksheet chứa cell</param>
        /// <returns>Tên hợp lệ hoặc null</returns>
        public static string GenerateValidRangeName(Range cell, Worksheet worksheet)
        {
            string baseName = $"Range_{worksheet?.Name ?? "Unknown"}_{cell?.Address[false, false]?.Replace("$", "").Replace(":", "_") ?? "Unknown"}";
            string validName = baseName;

            // Đảm bảo hợp lệ và unique
            int counter = 1;
            while (!IsValidNamedRangeName(validName) || NameExistsInWorkbook(validName, worksheet?.Application?.ActiveWorkbook))
            {
                validName = $"{baseName}_{counter}";
                counter++;

                // Tránh vòng lặp vô hạn
                if (counter > 1000)
                    break;
            }

            return IsValidNamedRangeName(validName) ? validName : null;
        }

        #endregion

        #region Active Objects Access

        /// <summary>
        /// Thử lấy các active objects (workbook và worksheet) hiện tại
        /// Kiểm tra an toàn để tránh exception khi Excel chưa sẵn sàng
        ///
        /// </summary>
        /// <param name="workbook">Output: Active workbook</param>
        /// <param name="worksheet">Output: Active worksheet</param>
        /// <returns>true nếu lấy thành công cả hai, false nếu thất bại</returns>
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
                // Silent fail - Excel có thể chưa sẵn sàng
                return false;
            }
        }

        #endregion
    }
}