using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace ExcelCustomAddin
{
    /// <summary>
    /// EvidenceCreationService - Quản lý việc tạo evidence sheets và hyperlinks
    ///
    /// Chức năng chính:
    /// - Tạo evidence sheets từ các cell được chọn
    /// - Xử lý merged cells và multiple cell selection
    /// - Tạo hyperlinks giữa source cells và evidence sheets
    /// - Định dạng evidence sheets theo cấu hình
    /// - Quản lý back buttons với named range references
    /// - Auto-generate sheet names cho cells trống
    ///
    /// Xử lý đặc biệt:
    /// - Merged cells: Chỉ xử lý first cell của merged area
    /// - Empty cells: Auto-generate names cho sheet "共通" và "テスト項目"
    /// - Existing sheets: Tạo hyperlink thay vì tạo mới
    /// - Named ranges: Sử dụng cho back button navigation
    ///
    /// Tác giả: lam.pt
    /// Ngày tạo: 2025
    /// </summary>
    public class EvidenceCreationService
    {
        #region Fields

        private readonly ThisAddIn _addIn;

        #endregion

        #region Constructor

        /// <summary>
        /// Khởi tạo EvidenceCreationService
        ///
        /// </summary>
        /// <param name="addIn">Instance của ThisAddIn chính</param>
        public EvidenceCreationService(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        #endregion

        #region Public Interface

        /// <summary>
        /// CreateEvidence - Phương thức chính để tạo evidence sheets
        ///
        /// Quy trình:
        /// 1. Validate inputs (workbook, sheet, selection)
        /// 2. Xử lý merged cells và multiple selection
        /// 3. Process từng cell để tạo sheet hoặc hyperlink
        /// 4. Hiển thị kết quả và cập nhật UI
        ///
        /// </summary>
        public void CreateEvidence()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var activeWorkbook = app.ActiveWorkbook;
                var activeSheet = app.ActiveSheet as Worksheet;

                // Validate inputs
                Range selectedRange;
                if (!ValidateEvidenceCreationInputs(activeWorkbook, activeSheet, app, out selectedRange))
                {
                    return;
                }

                // Get cells to process (handles both single and multiple cells, merged cells)
                var cellsToProcess = GetCellsToProcess(selectedRange);
                bool isMultipleCells = selectedRange.Cells.Count > 1;

                // Process each cell and create evidence sheets
                var (createdSheets, existingSheets, errorMessages) = ProcessEvidenceCells(
                    cellsToProcess, activeWorkbook, activeSheet);

                // Show results and update UI
                HandleEvidenceCreationResults(
                    isMultipleCells, createdSheets, existingSheets, errorMessages,
                    activeWorkbook, activeSheet);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi xảy ra khi tạo sheet bằng chứng: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region Input Validation

        /// <summary>
        /// Validate inputs cho việc tạo evidence
        /// Kiểm tra workbook, sheet, selection và protection status
        ///
        /// Validation checks:
        /// - Workbook phải đang mở
        /// - Sheet phải được chọn
        /// - Selection phải hợp lệ
        /// - Sheet không được bảo vệ
        ///
        /// </summary>
        /// <param name="activeWorkbook">Workbook hiện tại</param>
        /// <param name="activeSheet">Sheet hiện tại</param>
        /// <param name="app">Excel Application instance</param>
        /// <param name="selectedRange">Range được chọn (output parameter)</param>
        /// <returns>true nếu tất cả validation pass, false nếu có lỗi</returns>
        public bool ValidateEvidenceCreationInputs(Workbook activeWorkbook, Worksheet activeSheet, Microsoft.Office.Interop.Excel.Application app, out Range selectedRange)
        {
            selectedRange = null;

            if (activeWorkbook == null)
            {
                Logger.Error("Không có workbook nào đang mở trong CreateEvidence");
                MessageBox.Show("Không có workbook nào đang mở. Vui lòng mở một workbook và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (activeSheet == null)
            {
                Logger.Error("Không có sheet nào đang được chọn trong CreateEvidence");
                MessageBox.Show("Không có sheet nào đang được chọn. Vui lòng chọn một sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            try { selectedRange = app.Selection as Range; } catch { }
            if (selectedRange == null)
            {
                MessageBox.Show("Không có ô nào đang được chọn hoặc lựa chọn không hợp lệ. Vui lòng chọn một ô và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            if (activeSheet.ProtectContents || activeSheet.ProtectDrawingObjects || activeSheet.ProtectScenarios)
            {
                MessageBox.Show("Sheet đang được bảo vệ. Vui lòng bỏ bảo vệ sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        #endregion

        #region Cell Processing

        /// <summary>
        /// Lấy danh sách cells cần xử lý từ selected range
        /// Xử lý merged cells để tránh duplicate processing
        ///
        /// Logic xử lý:
        /// - Với multiple cells: Duyệt từng cell, skip merged cells đã xử lý
        /// - Với single cell: Xử lý merged area nếu có
        /// - Chỉ lấy first cell của mỗi merged area
        ///
        /// </summary>
        /// <param name="selectedRange">Range được chọn từ user</param>
        /// <returns>Danh sách cells cần xử lý (không duplicate merged cells)</returns>
        public List<Range> GetCellsToProcess(Range selectedRange)
        {
            var cellsToProcess = new List<Range>();
            var processedMergedAreas = new HashSet<string>();
            bool isMultipleCells = selectedRange.Cells.Count > 1;

            if (isMultipleCells)
            {
                foreach (Range cell in selectedRange.Cells)
                {
                    if (cell.MergeCells)
                    {
                        Range mergedArea = cell.MergeArea;
                        string mergedAreaAddress = mergedArea.Address[true, true];

                        if (!processedMergedAreas.Contains(mergedAreaAddress))
                        {
                            Range firstCell = mergedArea.Cells[1, 1];
                            cellsToProcess.Add(firstCell);
                            processedMergedAreas.Add(mergedAreaAddress);
                            Logger.Debug($"Merged cell detected at {cell.Address[false, false]}, using first cell {firstCell.Address[false, false]}");
                        }
                    }
                    else
                    {
                        cellsToProcess.Add(cell);
                    }
                }
            }
            else
            {
                if (selectedRange.MergeCells)
                {
                    Range mergedArea = selectedRange.MergeArea;
                    Range firstCell = mergedArea.Cells[1, 1];
                    cellsToProcess.Add(firstCell);
                    Logger.Debug($"Single merged cell detected at {selectedRange.Address[false, false]}, using first cell {firstCell.Address[false, false]}");
                }
                else
                {
                    cellsToProcess.Add(selectedRange);
                }
            }

            return cellsToProcess;
        }

        /// <summary>
        /// Lấy giá trị cell hoặc generate auto sheet name nếu cell trống
        /// Chỉ áp dụng auto-generation cho sheet "共通" và "テスト項目"
        ///
        /// </summary>
        /// <param name="cell">Cell cần lấy giá trị</param>
        /// <param name="activeSheet">Sheet chứa cell</param>
        /// <param name="cellsToProcessCount">Số lượng cells đang xử lý</param>
        /// <returns>Giá trị cell hoặc auto-generated name</returns>
        public string GetOrEvidenceNoText(Range cell, Worksheet activeSheet, int cellsToProcessCount)
        {
            string cellValue = cell.Value2 != null ? cell.Value2.ToString().Trim() : "";

            if (string.IsNullOrEmpty(cellValue) && cellsToProcessCount <= 1)
            {
                string currentSheetName = activeSheet.Name;
                if (currentSheetName == "共通" || currentSheetName == "テスト項目")
                {
                    cellValue = UtilityService.GenerateAutoEvidenceNoText(activeSheet, cell.Column, currentSheetName, 0);
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        cell.Value2 = cellValue;
                        Logger.Debug($"Auto-generated sheet name '{cellValue}' for cell {cell.Address[false, false]} (Column: {cell.Column})");
                    }
                }
            }

            return cellValue;
        }

        #endregion

        #region Evidence Processing

        /// <summary>
        /// Process các evidence cells và tạo/link sheets
        /// Xử lý từng cell để tạo sheet mới hoặc hyperlink đến sheet đã tồn tại
        ///
        /// Quy trình cho mỗi cell:
        /// 1. Lấy/generate cell value
        /// 2. Kiểm tra sheet đã tồn tại chưa
        /// 3. Tạo sheet mới hoặc hyperlink
        /// 4. Định dạng sheet mới nếu cần
        ///
        /// </summary>
        /// <param name="cellsToProcess">Danh sách cells cần xử lý</param>
        /// <param name="activeWorkbook">Workbook hiện tại</param>
        /// <param name="activeSheet">Sheet nguồn</param>
        /// <returns>Tuple chứa danh sách sheets đã tạo, sheets đã tồn tại, và error messages</returns>
        public (List<string> createdSheets, List<string> existingSheets, List<string> errorMessages)
            ProcessEvidenceCells(List<Range> cellsToProcess, Workbook activeWorkbook, Worksheet activeSheet, string evidenceNoText = "")
        {
            var createdSheets = new List<string>();
            var existingSheets = new List<string>();
            var errorMessages = new List<string>();

            // Build worksheet dictionary for faster lookup
            var worksheetDict = new Dictionary<string, Worksheet>(StringComparer.OrdinalIgnoreCase);
            foreach (Worksheet ws in activeWorkbook.Worksheets)
            {
                worksheetDict[ws.Name] = ws;
            }

            foreach (var cell in cellsToProcess)
            {
                try
                {
                    string cellValue = GetOrEvidenceNoText(cell, activeSheet, cellsToProcess.Count);

                    if (string.IsNullOrEmpty(cellValue))
                    {
                        errorMessages.Add($"Cell {cell.Address[false, false]} is empty and cannot generate auto sheet name");
                        continue;
                    }

                    if (worksheetDict.TryGetValue(cellValue, out Worksheet existingSheet))
                    {
                        // Sheet đã tồn tại, tạo hyperlink đến sheet đó
                        CreateHyperlinkToExistingSheet(cell, activeSheet, existingSheet, cellValue);
                        existingSheets.Add(cellValue);
                        Logger.Info($"Created hyperlink to existing sheet '{cellValue}' from cell {cell.Address[false, false]}");
                    }
                    else
                    {
                        // Tạo sheet mới
                        Worksheet newSheet = activeWorkbook.Worksheets.Add(After: activeWorkbook.Worksheets[activeWorkbook.Worksheets.Count]);
                        newSheet.Name = cellValue;

                        // Định dạng sheet mới
                        FormatNewEvidenceSheet(newSheet, cellValue);

                        // Tạo hyperlink đến sheet mới
                        CreateHyperlinkToNewSheet(cell, activeSheet, newSheet, cellValue);

                        createdSheets.Add(cellValue);
                        Logger.Info($"Created new sheet '{cellValue}' and hyperlink from cell {cell.Address[false, false]}");
                    }
                }
                catch (Exception ex)
                {
                    errorMessages.Add($"Error processing cell {cell.Address[false, false]}: {ex.Message}");
                    Logger.Error($"Error processing cell {cell.Address[false, false]}: {ex.Message}", ex);
                }
            }

            return (createdSheets, existingSheets, errorMessages);
        }

        #endregion

        #region Sheet Formatting

        /// <summary>
        /// Định dạng evidence sheet mới với các setting từ config
        /// Áp dụng page setup, margins, fonts, và view settings
        ///
        /// Các setting được áp dụng:
        /// - Print area và page orientation
        /// - Paper size và margins
        /// - Zoom/FitToPages (mutually exclusive)
        /// - Font và column widths
        /// - View mode và zoom
        ///
        /// </summary>
        /// <param name="newSheet">Sheet mới cần định dạng</param>
        /// <param name="sheetName">Tên sheet (để logging)</param>
        private void FormatNewEvidenceSheet(Worksheet newSheet, string sheetName)
        {
            try
            {
                var config = SheetConfigManager.GetGeneralConfig();

                // Set print area
                string printArea = $"A1:{SheetConfigManager.GetGeneralConfig().PageBreakColumnName}{config.PrintAreaLastRowIdx}";
                newSheet.PageSetup.PrintArea = printArea;

                // Set page setup
                if (config.PageOrientation.Equals("Landscape", StringComparison.OrdinalIgnoreCase))
                {
                    newSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                }
                else
                {
                    newSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait;
                }

                // Set paper size
                if (config.PaperSize.Equals("A4", StringComparison.OrdinalIgnoreCase))
                {
                    newSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4;
                }
                else if (config.PaperSize.Equals("A3", StringComparison.OrdinalIgnoreCase))
                {
                    newSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA3;
                }
                else if (config.PaperSize.Equals("Letter", StringComparison.OrdinalIgnoreCase))
                {
                    newSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperLetter;
                }
                else
                {
                    // Default to A4
                    newSheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA4;
                }

                // Set zoom or fit to pages (mutually exclusive in Excel)
                if (config.FitToPagesWide || config.FitToPagesTall)
                {
                    // If fit to pages is enabled, don't set zoom
                    newSheet.PageSetup.FitToPagesWide = config.FitToPagesWide ? 1 : 0;
                    newSheet.PageSetup.FitToPagesTall = config.FitToPagesTall ? 1 : 0;
                }
                else
                {
                    // If fit to pages is disabled, set zoom
                    newSheet.PageSetup.Zoom = config.Zoom;
                }

                // Set margins (convert from inches to points: 1 inch = 72 points)
                const double InchesToPoints = 72.0;
                newSheet.PageSetup.LeftMargin = config.LeftMargin * InchesToPoints;
                newSheet.PageSetup.RightMargin = config.RightMargin * InchesToPoints;
                newSheet.PageSetup.TopMargin = config.TopMargin * InchesToPoints;
                newSheet.PageSetup.BottomMargin = config.BottomMargin * InchesToPoints;
                newSheet.PageSetup.HeaderMargin = config.HeaderMargin * InchesToPoints;
                newSheet.PageSetup.FooterMargin = config.FooterMargin * InchesToPoints;
                newSheet.PageSetup.CenterHorizontally = config.CenterHorizontally;

                // Set column widths and row heights
                // ColumnWidth: measured in characters (default character width)
                // RowHeight: measured in points
                Range columnRange = newSheet.Range["A1", $"{SheetConfigManager.GetGeneralConfig().PageBreakColumnName}1"];
                columnRange.ColumnWidth = config.ColumnWidth;
                columnRange.RowHeight = config.RowHeight;

                // Set font
                newSheet.Cells.Font.Name = config.EvidenceFontName;
                newSheet.Cells.Font.Size = config.FontSize;

                // Set view mode
                newSheet.Activate();
                Globals.ThisAddIn.Application.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlPageBreakPreview;
                Globals.ThisAddIn.Application.ActiveWindow.Zoom = config.WindowZoom;

                Logger.Debug($"Formatted new evidence sheet '{sheetName}' with config settings");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error formatting new evidence sheet '{sheetName}': {ex.Message}");
            }
        }

        #endregion

        #region Hyperlink Management

        /// <summary>
        /// Tạo hyperlink đến existing sheet và setup back button
        /// Sử dụng named range cho back navigation khi có thể
        ///
        /// </summary>
        /// <param name="cell">Cell nguồn để tạo hyperlink</param>
        /// <param name="sourceSheet">Sheet chứa cell nguồn</param>
        /// <param name="existingSheet">Sheet đích đã tồn tại</param>
        /// <param name="sheetName">Tên sheet đích</param>
        public void CreateHyperlinkToExistingSheet(Range cell, Worksheet sourceSheet, Worksheet existingSheet, string sheetName)
        {
            // Tạo hyperlink đến sheet đã tồn tại
            sourceSheet.Hyperlinks.Add(cell, "", $"'{sheetName}'!A1", Type.Missing, sheetName);
            cell.Font.Name = SheetConfigManager.GetGeneralConfig().EvidenceFontName;

            // Tạo hoặc lấy named range cho cell gốc
            string namedRangeName = UtilityService.GetOrCreateNamedRangeForCell(cell, sourceSheet);

            // Setup back button
            int aColumnIndex = UtilityService.GetColumnIndex("A");
            Range backCell = existingSheet.Cells[1, aColumnIndex];
            backCell.Value2 = "<Back";

            try
            {
                if (backCell.Hyperlinks.Count > 0)
                {
                    backCell.Hyperlinks.Delete();
                }
            }
            catch { }

            // Tạo hyperlink back về cell gốc - sử dụng named range nếu có, ngược lại dùng địa chỉ trực tiếp
            string backAddress;
            if (!string.IsNullOrEmpty(namedRangeName))
            {
                backAddress = namedRangeName; // Sử dụng named range
                Logger.Debug($"Back button (existing sheet) sẽ reference đến named range: {namedRangeName}");
            }
            else
            {
                backAddress = $"'{sourceSheet.Name}'!{cell.Address[false, false]}"; // Fallback về địa chỉ trực tiếp
                Logger.Debug($"Back button (existing sheet) sẽ reference đến địa chỉ trực tiếp: {backAddress}");
            }

            existingSheet.Hyperlinks.Add(backCell, "", backAddress, Type.Missing, "<Back");
            backCell.Font.Name = SheetConfigManager.GetGeneralConfig().BackButtonFontName;
            backCell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
        }

        /// <summary>
        /// Tạo hyperlink đến new sheet và setup back button
        /// Sử dụng named range cho back navigation khi có thể
        ///
        /// </summary>
        /// <param name="cell">Cell nguồn để tạo hyperlink</param>
        /// <param name="sourceSheet">Sheet chứa cell nguồn</param>
        /// <param name="newSheet">Sheet đích mới tạo</param>
        /// <param name="sheetName">Tên sheet đích</param>
        public void CreateHyperlinkToNewSheet(Range cell, Worksheet sourceSheet, Worksheet newSheet, string sheetName)
        {
            // Tạo hyperlink đến sheet mới
            sourceSheet.Hyperlinks.Add(cell, "", $"'{sheetName}'!A1", Type.Missing, sheetName);
            cell.Font.Name = SheetConfigManager.GetGeneralConfig().EvidenceFontName;

            // Tạo hoặc lấy named range cho cell gốc
            string namedRangeName = UtilityService.GetOrCreateNamedRangeForCell(cell, sourceSheet);

            // Setup back button trên sheet mới
            int aColumnIndex = UtilityService.GetColumnIndex("A");
            Range backCell = newSheet.Cells[1, aColumnIndex];
            backCell.Value2 = "<Back";

            // Tạo hyperlink back về cell gốc - sử dụng named range nếu có, ngược lại dùng địa chỉ trực tiếp
            string backAddress;
            if (!string.IsNullOrEmpty(namedRangeName))
            {
                backAddress = namedRangeName; // Sử dụng named range
                Logger.Debug($"Back button (new sheet) sẽ reference đến named range: {namedRangeName}");
            }
            else
            {
                backAddress = $"'{sourceSheet.Name}'!{cell.Address[false, false]}"; // Fallback về địa chỉ trực tiếp
                Logger.Debug($"Back button (new sheet) sẽ reference đến địa chỉ trực tiếp: {backAddress}");
            }

            newSheet.Hyperlinks.Add(backCell, "", backAddress, Type.Missing, "<Back");
            backCell.Font.Name = SheetConfigManager.GetGeneralConfig().BackButtonFontName;
            backCell.HorizontalAlignment = XlHAlign.xlHAlignLeft;
        }

        #endregion

        #region Result Handling

        /// <summary>
        /// Xử lý kết quả tạo evidence và hiển thị messages
        /// Phân biệt xử lý cho single cell vs multiple cells
        ///
        /// </summary>
        /// <param name="isMultipleCells">Có phải multiple cells không</param>
        /// <param name="createdSheets">Danh sách sheets đã tạo</param>
        /// <param name="existingSheets">Danh sách sheets đã tồn tại</param>
        /// <param name="errorMessages">Danh sách error messages</param>
        /// <param name="activeWorkbook">Workbook hiện tại</param>
        /// <param name="activeSheet">Sheet hiện tại</param>
        public void HandleEvidenceCreationResults(
            bool isMultipleCells,
            List<string> createdSheets,
            List<string> existingSheets,
            List<string> errorMessages,
            Workbook activeWorkbook,
            Worksheet activeSheet)
        {
            if (isMultipleCells)
            {
                ShowMultipleCellsResults(createdSheets, existingSheets, errorMessages);
            }
            else
            {
                ShowSingleCellResults(createdSheets, existingSheets, errorMessages);
            }
        }

        /// <summary>
        /// Hiển thị kết quả cho multiple cells processing
        /// Hiển thị chi tiết số lượng sheets đã tạo, hyperlink, và errors
        ///
        /// </summary>
        /// <param name="createdSheets">Sheets đã tạo mới</param>
        /// <param name="existingSheets">Sheets đã tồn tại (hyperlink)</param>
        /// <param name="errorMessages">Error messages</param>
        private void ShowMultipleCellsResults(List<string> createdSheets, List<string> existingSheets, List<string> errorMessages)
        {
            var message = new System.Text.StringBuilder();

            if (existingSheets.Count > 0)
            {
                message.AppendLine($"Đã tạo hyperlink đến {existingSheets.Count} sheet đã tồn tại:");
                foreach (var sheet in existingSheets)
                    message.AppendLine($"  - {sheet}");
            }

            if (errorMessages.Count > 0)
            {
                message.AppendLine($"Có {errorMessages.Count} lỗi:");
                foreach (var error in errorMessages)
                    message.AppendLine($"  - {error}");
            }


            if (errorMessages.Count > 0 || existingSheets.Count > 1)
            {
                string title = errorMessages.Count > 0 ? "Hoàn thành với lỗi" : string.Empty;
                MessageBox.Show(message.ToString(), title, MessageBoxButton.OK,
                    errorMessages.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            }
        }

        /// <summary>
        /// Hiển thị kết quả cho single cell processing
        /// Chỉ hiển thị error nếu có, thành công thì im lặng
        ///
        /// </summary>
        /// <param name="createdSheets">Sheets đã tạo mới</param>
        /// <param name="existingSheets">Sheets đã tồn tại (hyperlink)</param>
        /// <param name="errorMessages">Error messages</param>
        private void ShowSingleCellResults(List<string> createdSheets, List<string> existingSheets, List<string> errorMessages)
        {
            if (errorMessages.Count > 0)
            {
                MessageBox.Show($"Có lỗi xảy ra: {errorMessages[0]}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        #endregion

        /// <summary>
        /// UpdateEvidenceNo
        /// </summary>
        public void UpdateEvidenceNo()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var activeWorkbook = app.ActiveWorkbook;
                var activeSheet = app.ActiveSheet as Worksheet;

                var sheetConfig = SheetConfigManager.GetSheetConfig(activeSheet.Name);
                string numberFormat = sheetConfig.NumberFormat;
                string prefix = sheetConfig.Prefix;
                bool.TryParse(sheetConfig.IsHorizontal, out bool isHorizontal);

                // Lấy dòng cuối cùng
                var lastRow =UtilityService.GetLastRow(activeSheet);

                if (isHorizontal)
                {
                    // Nếu là sheet ngang, thì chỉ cần thêm evidence mới vào cột cuối, nên không cần phải cập nhật
                    MessageBox.Show("Vì là sheet ngang, thì chỉ cần thêm evidence mới vào cột cuối, nên không cần phải cập nhật");
                    return;
                }
                else
                {
                    const string idColumnName = "B";
                    const string evidenceNoTextColumnName = "BH";
                    const string confirmColumnName = "AH";
                    const int evidenceNoStartRow = 3;

                    int evidenceNoTextColumnIdx = UtilityService.GetColumnIndex(evidenceNoTextColumnName);

                    for (int i = evidenceNoStartRow; i <= lastRow; i++)
                    {
                        var processingCell = activeSheet.Cells[i, evidenceNoTextColumnName];

                        // Lấy giá trị ID
                        var idCellValue = activeSheet.Cells[i, idColumnName].Value2;
                        if (idCellValue == null || string.IsNullOrWhiteSpace(idCellValue.ToString()))
                        {
                            Logger.Warning($"Dòng {i} không có giá trị ID trong cột {idColumnName}, bỏ qua cập nhật số bằng chứng.");
                            continue;
                        }
                        int.TryParse(idCellValue.ToString(), out int idCellValueInt);

                        // Lấy giá trị hiện tại của cell
                        string currentCellValue = activeSheet.Cells[i, evidenceNoTextColumnName].Value2?.ToString()?.Trim();
                        int currentIdInt = 0;

                        if (!string.IsNullOrEmpty(currentCellValue))
                        {
                            string[] currentValueSplitArr = currentCellValue.Split('_');
                            string currentIdStr = currentValueSplitArr.Length > 1 ? currentValueSplitArr[currentValueSplitArr.Length - 1] : "0";
                            int.TryParse(currentIdStr, out currentIdInt);
                        }

                        // Lấy ra EvidenceText: vd: "エビデンス_001"
                        string evidenceNoText = UtilityService.GenerateAutoEvidenceNoText(activeSheet, evidenceNoTextColumnIdx, activeSheet.Name, idCellValueInt);

                        // Xoá Named Range theo newEvidenceNo nếu đã tồn tại
                        DeleteNamedRangeByName(activeWorkbook, evidenceNoText);

                        // Kiểm tra xem cột Kết quả mong muốn có giá trị hay không?
                        var confirmResultCell = activeSheet.Cells[i, confirmColumnName];
                        var confirmResultCellValue = confirmResultCell.Value2;
                        if (confirmResultCellValue == null || string.IsNullOrWhiteSpace(confirmResultCellValue.ToString()))
                        {
                            // Nếu không có giá trị thì không cần xử lý tiếp theo
                            continue;
                        }

                        if (idCellValueInt > 0)
                        {
                            // Nếu cell hiện tại đang trống
                            if (string.IsNullOrWhiteSpace(currentCellValue))
                            {
                                // Thực hiện backup sheet nếu sheet với tên mới đã tồn tại
                                bool sheetExists = IsSheetExist(activeWorkbook, evidenceNoText);
                                if (sheetExists)
                                {
                                    // Nếu sheet đã tồn tại, thực hiện đổi tên để backup
                                    string backupSheetName = $"{evidenceNoText}_Backup";
                                    activeWorkbook.Worksheets[evidenceNoText].Name = backupSheetName;
                                }

                                // Setting newEvidenceNo cho cell hiện tại
                                activeSheet.Cells[i, evidenceNoTextColumnName].Value2 = evidenceNoText;

                                // Tạo sheet tương ứng
                                var cellsToProcess = new List<Range> { activeSheet.Cells[i, evidenceNoTextColumnName] };
                                ProcessEvidenceCells(cellsToProcess, activeWorkbook, activeSheet, evidenceNoText);

                                // Logger thông tin
                                Logger.Info($"Cập nhật số bằng chứng cho dòng {i}: {evidenceNoText}");

                                continue;
                            }
                            else
                            {
                                // Nếu cell hiện tại có giá trị, so sánh với ID
                                if (currentIdInt == idCellValueInt)
                                {
                                    // Nếu số hiện tại đã đúng, không cần cập nhật
                                    continue;
                                }
                                else
                                {
                                    // Nếu số hiện tại khác với ID
                                    // Cập nhật số bằng chứng mới
                                    activeSheet.Cells[i, evidenceNoTextColumnName].Value2 = evidenceNoText;

                                    // Kiểm tra xem sheet với tên mới có tồn tại không
                                    bool newSheetExists = IsSheetExist(activeWorkbook, evidenceNoText);

                                    if (newSheetExists)
                                    {
                                        // Nếu sheet đã tồn tại, thực hiện đổi tên để backup
                                        string backupSheetName = $"{evidenceNoText}_Backup";
                                        activeWorkbook.Worksheets[evidenceNoText].Name = backupSheetName;
                                    }

                                    // Kiểm tra xem sheet với số bằng chứng hiện tại + chuỗi "_Backup" có tồn tại không
                                    bool backupSheetExists = IsSheetExist(activeWorkbook, $"{currentCellValue}_Backup");
                                    if (backupSheetExists)
                                    {
                                        // Nếu backup sheet tồn tại 
                                        // Đổi tên sheet có tên _Backup thành số bằng chứng mới
                                        activeWorkbook.Worksheets[$"{currentCellValue}_Backup"].Name = evidenceNoText;

                                        // Tạo hyperlink đến sheet mới
                                        CreateHyperlinkToNewSheet(processingCell, activeSheet, activeWorkbook.Worksheets[evidenceNoText], evidenceNoText);
                                    }
                                    else
                                    {
                                        // Nếu backup sheet không tồn tại
                                        // Gọi hàm ProcessEvidenceCells để tạo sheet mới với số bằng chứng mới
                                        var cellsToProcess = new List<Range> { activeSheet.Cells[i, evidenceNoTextColumnName] };
                                        ProcessEvidenceCells(cellsToProcess, activeWorkbook, activeSheet, evidenceNoText);
                                    }
                                }
                            }
                        }
                    }
                }

                MessageBox.Show("Cập nhật số bằng chứng hoàn tất.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Có lỗi xảy ra khi cập nhật số bằng chứng: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool IsSheetExist(Workbook workbook, string sheetName)
        {
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Xóa Named Range theo tên
        /// Duyệt qua tất cả named ranges trong workbook và xóa các range có tên khớp
        /// 
        /// </summary>
        /// <param name="workbook">Workbook chứa các named ranges</param>
        /// <param name="rangeName">Tên của named range cần xóa</param>
        private void DeleteNamedRangeByName(Workbook workbook, string rangeName)
        {
            if (workbook == null || string.IsNullOrEmpty(rangeName))
                return;

            try
            {
                var names = workbook.Names;
                // Duyệt ngược để tránh lỗi index khi xóa
                for (int i = names.Count; i >= 1; i--)
                {
                    try
                    {
                        var nameObj = names.Item(i);
                        if (nameObj.Name.Equals(rangeName, StringComparison.OrdinalIgnoreCase))
                        {
                            nameObj.Delete();
                            Logger.Debug($"Xóa Named Range: {rangeName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"Lỗi khi xóa named range tại index {i}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Lỗi khi xóa named range '{rangeName}': {ex.Message}");
            }
        }
    }
}