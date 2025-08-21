/**
 * Worksheet operations utilities
 */

import ExcelUtils from './ExcelUtils';
import StorageService from '../services/StorageService';

export class WorksheetOperations {
  /**
   * Tạo Evidence Sheet từ giá trị cell hiện tại
   */
  static async createEvidenceSheet(): Promise<{ success: boolean; message: string }> {
    try {
      const cellInfo = await ExcelUtils.getSelectedCellInfo();
      
      if (!cellInfo.value || cellInfo.value.toString().trim() === "") {
        return {
          success: false,
          message: "Ô hiện tại đang để trống. Vui lòng nhập giá trị vào ô và thử lại."
        };
      }

      const newSheetName = cellInfo.value.toString().trim();
      
      // Kiểm tra sheet đã tồn tại chưa
      const exists = await ExcelUtils.worksheetExists(newSheetName);
      if (exists) {
        // Tạo hyperlink đến sheet đã tồn tại
        await ExcelUtils.createHyperlink(
          cellInfo.address,
          `#'${newSheetName}'!A1`,
          newSheetName
        );
        return {
          success: true,
          message: `Sheet '${newSheetName}' đã tồn tại. Đã tạo hyperlink.`
        };
      }

      // Tạo sheet mới
      return await Excel.run(async (context) => {
        const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
        activeWorksheet.load("name");
        
        const newWorksheet = context.workbook.worksheets.add(newSheetName);
        
        await context.sync();
        
        // Thiết lập format cho sheet mới - thực hiện trong cùng context
        const formatRange = newWorksheet.getRange("A1:BC48");
        formatRange.format.font.name = "MS PGothic";
        formatRange.format.font.size = 9;
        
        // Thiết lập row height cho toàn bộ range
        formatRange.format.rowHeight = 12.75;
        
        // Thiết lập column width cho toàn bộ sheet
        const allColumns = newWorksheet.getRange("1:1").getEntireColumn();
        allColumns.format.columnWidth = 14;

        // Thiết lập page layout
        try {
          newWorksheet.pageLayout.orientation = Excel.PageOrientation.landscape;
          const printRange = newWorksheet.getRange("A1:BC48");
          newWorksheet.pageLayout.setPrintArea(printRange);
          
          // Thiết lập Fit to Page (tương đương với setFitToPage(true) trong PHP)
          // Trong Excel JavaScript API, chúng ta sử dụng zoom để fit to page
          newWorksheet.pageLayout.zoom = {
            horizontalFitToPages: 1,
            verticalFitToPages: 1
          };
        } catch (layoutError) {
          console.warn("Page layout settings not fully supported:", layoutError);
        }

        // Lấy address chính xác cho hyperlink back (loại bỏ sheet name nếu có)
        const cleanAddress = cellInfo.address.includes('!') 
          ? cellInfo.address.split('!')[1] 
          : cellInfo.address;

        // Tạo hyperlink từ ô hiện tại đến sheet mới
        const originalCell = activeWorksheet.getRange(cleanAddress);
        originalCell.hyperlink = {
          address: `#'${newSheetName}'!A1`,
          textToDisplay: newSheetName
        };

        // Đặt "Back" vào A1 và tạo hyperlink back
        const backCell = newWorksheet.getRange("A1");
        backCell.values = [["Back"]];
        backCell.hyperlink = {
          address: `#'${activeWorksheet.name}'!${cleanAddress}`,
          textToDisplay: "戻る"
        };
        
        await context.sync();
        
        return {
          success: true,
          message: `Đã tạo sheet '${newSheetName}' thành công!`
        };
      });
    } catch (error) {
      console.error("Error creating evidence sheet:", error);
      return {
        success: false,
        message: `Có lỗi xảy ra: ${error.message}`
      };
    }
  }

  /**
   * Format tất cả worksheets (focus A1, zoom 100%)
   */
  static async formatDocument(): Promise<{ success: boolean; message: string }> {
    try {
      return await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items");
        await context.sync();

        // Duyệt qua tất cả worksheets
        for (const worksheet of worksheets.items) {
          worksheet.activate();
          await context.sync();
          
          // Focus A1
          worksheet.getRange("A1").select();
          await context.sync();
        }

        // Kích hoạt lại worksheet đầu tiên
        if (worksheets.items.length > 0) {
          worksheets.items[0].activate();
          await context.sync();
        }

        return {
          success: true,
          message: "Đã format document thành công!"
        };
      });
    } catch (error) {
      console.error("Error formatting document:", error);
      return {
        success: false,
        message: `Có lỗi xảy ra: ${error.message}`
      };
    }
  }

  /**
   * Đổi tên worksheet với validation
   */
  static async renameWorksheet(oldName: string, newName: string): Promise<{ success: boolean; message: string }> {
    try {
      const trimmedName = newName.trim();

      // Validation tên sheet
      if (trimmedName.length > 31) {
        return {
          success: false,
          message: "Tên sheet không được vượt quá 31 ký tự."
        };
      }

      // Kiểm tra ký tự không hợp lệ
      const invalidChars = ['\\', '/', '?', '*', '[', ']', ':'];
      const hasInvalidChars = invalidChars.some(char => trimmedName.includes(char));
      if (hasInvalidChars) {
        return {
          success: false,
          message: "Tên sheet không được chứa các ký tự: \\ / ? * [ ] :"
        };
      }

      return await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name");
        const worksheets = workbook.worksheets;
        worksheets.load("items/name");
        
        const targetWorksheet = worksheets.getItem(oldName);
        targetWorksheet.load("name,protection");
        
        await context.sync();

        // Kiểm tra sheet có bị bảo vệ không
        if (targetWorksheet.protection.protected) {
          return {
            success: false,
            message: `Sheet '${oldName}' đang được bảo vệ. Vui lòng bỏ bảo vệ sheet và thử lại.`
          };
        }

        // Kiểm tra tên sheet đã tồn tại chưa
        const existingNames = worksheets.items.map(ws => ws.name.toLowerCase());
        if (existingNames.includes(trimmedName.toLowerCase()) && trimmedName !== oldName) {
          return {
            success: false,
            message: `Sheet có tên '${trimmedName}' đã tồn tại. Vui lòng chọn tên khác.`
          };
        }

        targetWorksheet.name = trimmedName;
        await context.sync();
        
        // Cập nhật pinned sheets storage nếu sheet này được pin
        const workbookName = workbook.name;
        if (StorageService.isSheetPinned(workbookName, oldName)) {
          StorageService.togglePinSheet(workbookName, oldName); // Remove old
          StorageService.togglePinSheet(workbookName, trimmedName); // Add new
        }
        
        return {
          success: true,
          message: `Đã đổi tên sheet từ '${oldName}' thành '${trimmedName}' thành công!`
        };
      });
    } catch (error) {
      console.error("Error changing sheet name:", error);
      return {
        success: false,
        message: `Có lỗi xảy ra: ${error.message}`
      };
    }
  }

  /**
   * Tạo evidence sheet  /**
   * Debug method để kiểm tra hyperlinks trong workbook
   */
  static async debugHyperlinks(): Promise<{ success: boolean; message: string }> {
    try {
      return await Excel.run(async (context) => {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();

        let hyperlinksFound = 0;
        let debugInfo = "=== DEBUG HYPERLINKS ===\n";

        for (const worksheet of worksheets.items) {
          debugInfo += `\nWorksheet: ${worksheet.name}\n`;
          
          try {
            const usedRange = worksheet.getUsedRangeOrNullObject();
            usedRange.load("rowCount, columnCount, address");
            await context.sync();

            if (!usedRange.isNullObject && usedRange.rowCount && usedRange.columnCount) {
              const maxRows = Math.min(usedRange.rowCount, 100);
              const maxCols = Math.min(usedRange.columnCount, 50);
              
              for (let row = 0; row < maxRows; row++) {
                for (let col = 0; col < maxCols; col++) {
                  try {
                    const cell = usedRange.getCell(row, col);
                    cell.load("hyperlink, address");
                    await context.sync();

                    if (cell.hyperlink && cell.hyperlink.address) {
                      hyperlinksFound++;
                      debugInfo += `  Cell ${cell.address}: ${cell.hyperlink.address} -> "${cell.hyperlink.textToDisplay}"\n`;
                    }
                  } catch (cellError) {
                    continue;
                  }
                }
              }
            }
          } catch (worksheetError) {
            debugInfo += `  Error reading worksheet: ${worksheetError}\n`;
          }
        }

        debugInfo += `\nTotal hyperlinks found: ${hyperlinksFound}`;
        console.log(debugInfo);

        return {
          success: true,
          message: `Found ${hyperlinksFound} hyperlinks. Check console for details.`
        };
      });
    } catch (error) {
      console.error("Error debugging hyperlinks:", error);
      return {
        success: false,
        message: `Error: ${error.message}`
      };
    }
  }
}

export default WorksheetOperations;
