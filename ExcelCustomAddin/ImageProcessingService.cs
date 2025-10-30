using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace ExcelCustomAddin
{
  /// <summary>
  /// Service xử lý các chức năng liên quan đến hình ảnh
  /// </summary>
  public class ImageProcessingService
  {
    private readonly ThisAddIn _addIn;

    // Column used as the page-break / right-most printed column for Evidence sheets
    private static string PAGE_BREAK_COLUMN_NAME => SheetConfigManager.GetGeneralConfig().PageBreakColumnName;
    // Last row index for print area in Evidence sheets
    private static int PRINT_AREA_LAST_ROW_IDX => SheetConfigManager.GetGeneralConfig().PrintAreaLastRowIdx;

    public ImageProcessingService(ThisAddIn addIn)
    {
      _addIn = addIn;
    }

    /// <summary>
    /// Apply picture style to shape: no fill, no line, shadow, no reflection, no glow
    /// </summary>
    /// <param name="shape">The shape to apply style to</param>
    public void ApplyPictureStyleToShape(Microsoft.Office.Interop.Excel.Shape shape)
    {
      shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
      shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
      shape.Shadow.Style = Microsoft.Office.Core.MsoShadowStyle.msoShadowStyleOuterShadow;
      shape.Shadow.Type = Microsoft.Office.Core.MsoShadowType.msoShadow21;
      shape.Shadow.ForeColor.RGB = 0; // Black
      shape.Shadow.Transparency = 0.3f;
      shape.Shadow.Size = 100;
      shape.Shadow.Blur = 15;
      shape.Shadow.OffsetX = 0;
      shape.Shadow.OffsetY = 0;
      shape.Reflection.Type = Microsoft.Office.Core.MsoReflectionType.msoReflectionTypeNone;
      shape.Glow.Radius = 0;
      shape.SoftEdge.Radius = 0;
    }

    /// <summary>
    /// Tính toán tỷ lệ scale tự động dựa trên print area
    /// </summary>
    /// <param name="worksheet">Worksheet chứa hình ảnh</param>
    /// <param name="imageWidth">Chiều rộng gốc của hình ảnh</param>
    /// <returns>Tỷ lệ scale (0.0 - 1.0)</returns>
    public double CalculateAutoScaleRate(Worksheet worksheet, double imageWidth)
    {
      try
      {
        // Lấy chiều rộng của print area (từ cột A đến PAGE_BREAK_COLUMN_NAME)
        int startColumn = UtilityService.GetColumnIndex("A");
        int endColumn = UtilityService.GetColumnIndex(PAGE_BREAK_COLUMN_NAME);

        double printAreaWidth = 0;
        for (int col = startColumn; col <= endColumn; col++)
        {
          Range columnRange = worksheet.Cells[1, col];
          printAreaWidth += (double)columnRange.Width;
        }

        // Trừ đi chiều rộng của 2 cột (margin trái và phải)
        double availableWidth = printAreaWidth - (2 * (double)worksheet.Cells[1, 1].Width);

        // Tính tỷ lệ scale
        double scaleRate = availableWidth / imageWidth;

        // Giới hạn tỷ lệ trong khoảng hợp lý (10% - 100%)
        scaleRate = Math.Max(0.1, Math.Min(1.0, scaleRate));

        Logger.Debug($"Auto scale calculation: PrintAreaWidth={printAreaWidth:F1}, AvailableWidth={availableWidth:F1}, ImageWidth={imageWidth:F1}, ScaleRate={scaleRate:F3}");

        return scaleRate;
      }
      catch (Exception ex)
      {
        Logger.Warning($"Error calculating auto scale rate: {ex.Message}, using default 1.0");
        return 1.0;
      }
    }

    /// <summary>
    /// Lấy tỷ lệ scale dựa trên cài đặt auto fix width
    /// </summary>
    /// <param name="worksheet">Worksheet hiện tại</param>
    /// <param name="imageWidth">Chiều rộng gốc của hình ảnh (đối với auto scale)</param>
    /// <returns>Tỷ lệ scale</returns>
    public double GetScaleRate(Worksheet worksheet, double imageWidth = 0)
    {
      if (_addIn._actionPanel.cbAutoFixWidth.Checked)
      {
        // Tự động tính toán tỷ lệ dựa trên print area
        return CalculateAutoScaleRate(worksheet, imageWidth);
      }
      else
      {
        // Sử dụng tỷ lệ từ numScalePercent
        return (double)_addIn._actionPanel.numScalePercent.Value / 100.0;
      }
    }

    /// <summary>
    /// FormatImages - Xử lý format hình ảnh đã chọn
    /// </summary>
    public void FormatImages()
    {
      try
      {
        var app = Globals.ThisAddIn.Application;
        var activeSheet = app.ActiveSheet as Worksheet;

        // Lấy các shape đã chọn
        ShapeRange selectedShapes = null;
        try
        {
          selectedShapes = app.Selection.ShapeRange;
        }
        catch
        {
          Logger.Error("Không có hình ảnh nào được chọn. Vui lòng chọn các hình ảnh và thử lại.");
          MessageBox.Show("Không có hình ảnh nào được chọn. Vui lòng chọn các hình ảnh và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
          return;
        }

        if (selectedShapes == null || selectedShapes.Count == 0)
        {
          Logger.Error("Không có hình ảnh nào được chọn.");
          MessageBox.Show("Không có hình ảnh nào được chọn.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
          return;
        }

        // Lấy tỷ lệ scale từ action panel hoặc tự động tính toán
        int formattedCount = 0;
        foreach (Shape shape in selectedShapes)
        {
          try
          {
            // Lấy tỷ lệ scale cho hình ảnh này
            double scalePercent = GetScaleRate(activeSheet, shape.Width);

            // 1. Scale theo tỷ lệ đã tính
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            shape.Height = (float)(shape.Height * scalePercent);

            // 2. Apply picture style
            ApplyPictureStyleToShape(shape);

            formattedCount++;
          }
          catch (Exception ex)
          {
            Logger.Warning($"Lỗi khi format hình ảnh '{shape.Name}': {ex.Message}");
          }
        }

        Logger.Info($"Đã format {formattedCount} hình ảnh thành công.");
      }
      catch (Exception ex)
      {
        Logger.Error($"Có lỗi xảy ra khi format hình ảnh: {ex.Message}", ex);
        MessageBox.Show($"Có lỗi xảy ra khi format hình ảnh: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
      }
    }

    /// <summary>
    /// Tính toán vị trí ban đầu để chèn hình ảnh
    /// </summary>
    public (double Top, double Left) CalculateInitialImagePosition(Worksheet activeSheet, Range activeCell, bool insertOnNewPage)
    {
      // Luôn chèn hình ảnh đầu tiên tại chính cell đang chọn
      double topLocation = (double)activeCell.Top;
      double leftLocation = (double)activeCell.Left;

      Logger.Debug($"Insert first image at cell {activeCell.Address[false, false]} (Row: {activeCell.Row}, Column: {activeCell.Column})");

      return (topLocation, leftLocation);
    }

    /// <summary>
    /// Thêm page break cho hình ảnh tiếp theo nếu cần thiết
    /// </summary>
    public void AddPageBreakForNextImage(Worksheet activeSheet, Range activeCell, ref (double Top, double Left) position, ref double currentPageStartTop, double maxBottomPosition, int imageIndex, double nextImageHeight)
    {
      try
      {
        // Tính toán chiều cao của trang hiện tại
        double pageHeight = activeCell.Height * PRINT_AREA_LAST_ROW_IDX;

        // Tính toán chiều cao đã sử dụng trên trang hiện tại
        double usedHeightOnCurrentPage = maxBottomPosition - currentPageStartTop;

        // Tính toán chiều cao sẽ sử dụng nếu chèn hình tiếp theo (bao gồm khoảng cách giữa các hình)
        double nextImageHeightWithSpacing = nextImageHeight + activeCell.Height;

        // Kiểm tra xem có cần tạo page break không
        // Tạo page break nếu: chiều cao đã dùng + chiều cao hình tiếp theo > chiều cao trang hiện tại
        if (usedHeightOnCurrentPage + nextImageHeightWithSpacing > pageHeight)
        {
          // Chèn horizontal page break
          int rowForPageBreak = Math.Max(1, (int)Math.Ceiling(maxBottomPosition / activeCell.Height) + 2);
          Range breakRange = activeSheet.Cells[rowForPageBreak, 1];
          activeSheet.HPageBreaks.Add(breakRange);

          // Cập nhật vị trí cho ảnh mới tại row 2 của trang mới (cùng cột với cell ban đầu)
          Range row2Cell = activeSheet.Cells[rowForPageBreak + 1, activeCell.Column];
          position.Top = (double)row2Cell.Top;
          position.Left = (double)row2Cell.Left;

          // Cập nhật vị trí bắt đầu của trang mới
          currentPageStartTop = position.Top;

          Logger.Debug($"Added page break at row {rowForPageBreak}, image {imageIndex} will be placed at row {rowForPageBreak + 1} (row 2 of new page). Used height: {usedHeightOnCurrentPage:F1}, Next image height: {nextImageHeight:F1}, Total would be: {usedHeightOnCurrentPage + nextImageHeightWithSpacing:F1}, Page height: {pageHeight:F1}");
        }
        else
        {
          Logger.Debug($"No page break needed for image {imageIndex}. Used height: {usedHeightOnCurrentPage:F1}, Next image height: {nextImageHeight:F1}, Total: {usedHeightOnCurrentPage + nextImageHeightWithSpacing:F1} / Page height: {pageHeight:F1}. Images will fit on current page.");
        }
      }
      catch (Exception pageBreakEx)
      {
        Logger.Warning($"Failed to add page break: {pageBreakEx.Message}");
        // Tiếp tục chèn ảnh mà không có page break
      }
    }

    /// <summary>
    /// Cập nhật vị trí cho hình ảnh tiếp theo
    /// </summary>
    public void UpdatePositionForNextImage(Worksheet activeSheet, Range activeCell, ref (double Top, double Left) position, Microsoft.Office.Interop.Excel.Shape shape, bool insertOnNewPage)
    {
      // Luôn cập nhật vị trí cho hình ảnh tiếp theo (xếp theo chiều dọc)
      // Page break sẽ được xử lý riêng trong AddPageBreakForNextImage
      position.Top += shape.Height + activeCell.Height;

      Logger.Debug($"Updated position for next image: Top={position.Top:F1}, Left={position.Left:F1}");
    }

    /// <summary>
    /// Điều chỉnh print area sau khi chèn hình ảnh
    /// </summary>
    public void AdjustPrintAreaForImages(Worksheet activeSheet, Range activeCell, double maxBottomPosition)
    {
      try
      {
        // Tính toán row cuối cùng cần thiết cho print area
        int lastRow = Math.Max(PRINT_AREA_LAST_ROW_IDX, (int)Math.Ceiling(maxBottomPosition / activeCell.Height) + 2);

        // Cập nhật print area
        string printArea = $"'{activeSheet.Name}'!$A$1:${PAGE_BREAK_COLUMN_NAME}${lastRow}";
        activeSheet.PageSetup.PrintArea = printArea;

        Logger.Debug($"Adjusted print area to: {printArea}");
      }
      catch (Exception ex)
      {
        Logger.Warning($"Failed to adjust print area: {ex.Message}");
      }
    }

    /// <summary>
    /// Xử lý sự kiện chèn nhiều hình ảnh
    /// </summary>
    public void InsertMultipleImages()
    {
      try
      {
        var app = Globals.ThisAddIn.Application;
        var activeWorkbook = app.ActiveWorkbook;
        var activeSheet = app.ActiveSheet as Worksheet;

        // Kiểm tra workbook và sheet
        if (activeWorkbook == null)
        {
          Logger.Error("Không có workbook nào đang mở. Vui lòng mở một workbook và thử lại.");
          MessageBox.Show("Không có workbook nào đang mở. Vui lòng mở một workbook và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
          return;
        }
        if (activeSheet == null)
        {
          Logger.Error("Không có sheet nào đang được chọn. Vui lòng chọn một sheet và thử lại.");
          MessageBox.Show("Không có sheet nào đang được chọn. Vui lòng chọn một sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
          return;
        }

        // Kiểm tra cell đang chọn
        Range activeCell = null;
        try { activeCell = app.ActiveCell as Range; } catch { }
        if (activeCell == null)
        {
          Logger.Error("Không có ô nào đang được chọn hoặc lựa chọn không hợp lệ. Vui lòng chọn một ô và thử lại.");
          MessageBox.Show("Không có ô nào đang được chọn hoặc lựa chọn không hợp lệ. Vui lòng chọn một ô và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
          return;
        }

        // Kiểm tra sheet có bị bảo vệ không
        if (activeSheet.ProtectContents || activeSheet.ProtectDrawingObjects || activeSheet.ProtectScenarios)
        {
          Logger.Error("Sheet đang được bảo vệ. Vui lòng bỏ bảo vệ sheet và thử lại.");
          MessageBox.Show("Sheet đang được bảo vệ. Vui lòng bỏ bảo vệ sheet và thử lại.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
          return;
        }

        // Đường dẫn thư mục chứa hình ảnh
        string folderPath = _addIn._actionPanel.txtImagePath.Text.Trim();

        // Tạo thư mục nếu chưa tồn tại
        if (!System.IO.Directory.Exists(folderPath))
        {
          System.IO.Directory.CreateDirectory(folderPath);
          Logger.Info($"Đã tạo thư mục '{folderPath}'. Vui lòng thêm hình ảnh vào thư mục này và thử lại.");
          MessageBox.Show($"Đã tạo thư mục '{folderPath}'. Vui lòng thêm hình ảnh vào thư mục này và thử lại.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
          return;
        }

        // Lấy danh sách file hình ảnh
        string[] imageExtensions = { "*.jpg", "*.jpeg", "*.png", "*.bmp", "*.gif", "*.tiff" };
        var imageFiles = new List<string>();

        foreach (string extension in imageExtensions)
        {
          var files = System.IO.Directory.GetFiles(folderPath, extension, System.IO.SearchOption.TopDirectoryOnly);
          imageFiles.AddRange(files);
        }

        if (imageFiles.Count == 0)
        {
          Logger.Warning($"Không tìm thấy file hình ảnh nào trong thư mục '{folderPath}'. Các định dạng được hỗ trợ: JPG, JPEG, PNG, BMP, GIF, TIFF");
          MessageBox.Show($"Không tìm thấy file hình ảnh nào trong thư mục '{folderPath}'.\nCác định dạng được hỗ trợ: JPG, JPEG, PNG, BMP, GIF, TIFF", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
          return;
        }

        // Khởi tạo các biến cần thiết
        bool insertOnNewPage = _addIn._actionPanel.chkInsertOnNewPage.Checked;
        Logger.Info($"Insert on new page mode: {insertOnNewPage}");

        // Tính toán vị trí bắt đầu
        var position = CalculateInitialImagePosition(activeSheet, activeCell, insertOnNewPage);
        int insertedCount = 0;
        int errorCount = 0;
        double maxBottomPosition = position.Top;
        double currentPageStartTop = position.Top; // Theo dõi vị trí bắt đầu của trang hiện tại

        // Chèn từng hình ảnh
        foreach (string imagePath in imageFiles)
        {
          try
          {
            // Bước 1: Tải hình ảnh tạm thời để lấy kích thước gốc
            var tempShape = activeSheet.Shapes.AddPicture(
                imagePath,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue,
                0, 0, -1, -1
            );
            tempShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

            // Lấy tỷ lệ scale cho hình ảnh này
            double resizeRate = GetScaleRate(activeSheet, tempShape.Width);

            // Áp dụng tỷ lệ scale để lấy kích thước sau khi resize
            tempShape.Height = (float)(tempShape.Height * resizeRate);
            double imageHeight = tempShape.Height;
            tempShape.Delete(); // Xóa hình tạm

            // Bước 2: Kiểm tra và thêm page break nếu cần (dựa vào chiều cao hình sắp chèn)
            if (insertOnNewPage && insertedCount > 0)
            {
              AddPageBreakForNextImage(activeSheet, activeCell, ref position, ref currentPageStartTop, maxBottomPosition, insertedCount + 1, imageHeight);
            }

            // Bước 3: Chèn hình ảnh vào vị trí đã tính toán
            var shape = activeSheet.Shapes.AddPicture(
                imagePath,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue,
                (float)position.Left,
                (float)position.Top,
                -1, // Width - tự động
                -1  // Height - tự động
            );

            // Điều chỉnh kích thước hình ảnh
            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            shape.Height = (float)(shape.Height * resizeRate);

            // Áp dụng picture style
            ApplyPictureStyleToShape(shape);

            // Cập nhật vị trí cho hình ảnh tiếp theo
            UpdatePositionForNextImage(activeSheet, activeCell, ref position, shape, insertOnNewPage);

            // Theo dõi vị trí thấp nhất
            maxBottomPosition = Math.Max(maxBottomPosition, shape.Top + shape.Height);

            insertedCount++;

            // Xóa file sau khi chèn thành công
            try
            {
              // Xóa thuộc tính readonly nếu có
              System.IO.File.SetAttributes(imagePath, System.IO.FileAttributes.Normal);
              System.IO.File.Delete(imagePath);
            }
            catch (Exception deleteEx)
            {
              Logger.Warning($"Không thể xóa file {imagePath}: {deleteEx.Message}");
            }
          }
          catch (Exception ex)
          {
            errorCount++;
            Logger.Error($"Lỗi khi chèn hình ảnh {imagePath}: {ex.Message}");
          }
        }

        // Tự động điều chỉnh print area sau khi chèn hình ảnh
        if (insertedCount > 0)
        {
          AdjustPrintAreaForImages(activeSheet, activeCell, maxBottomPosition);
          Logger.Info($"Print area adjusted for {insertedCount} inserted images");
        }

        Logger.Info($"Successfully inserted {insertedCount} images, {errorCount} errors");
      }
      catch (Exception ex)
      {
        Logger.Error($"Có lỗi xảy ra khi chèn hình ảnh: {ex.Message}", ex);
        MessageBox.Show($"Có lỗi xảy ra khi chèn hình ảnh: {ex.Message}", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
      }
    }
  }
}