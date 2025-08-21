/**
 * Utility functions for Excel operations
 */

export interface SheetInfo {
  name: string;
  tabColor?: string;
  isPinned: boolean;
}

export class ExcelUtils {
  /**
   * Lấy thông tin workbook hiện tại
   */
  static async getWorkbookInfo(): Promise<{ name: string; path: string }> {
    return Excel.run(async (context) => {
      const workbook = context.workbook;
      workbook.load("name");
      await context.sync();
      
      return {
        name: workbook.name || "Untitled Workbook",
        path: workbook.name || "Chưa lưu workbook"
      };
    });
  }

  /**
   * Lấy danh sách tất cả worksheets
   */
  static async getWorksheetList(): Promise<{ worksheets: SheetInfo[]; activeSheet: string; workbookName: string }> {
    return Excel.run(async (context) => {
      const workbook = context.workbook;
      workbook.load("name");
      
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name,items/tabColor");
      
      const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
      activeWorksheet.load("name");
      
      await context.sync();

      const worksheetList: SheetInfo[] = worksheets.items.map(sheet => ({
        name: sheet.name,
        tabColor: sheet.tabColor || undefined,
        isPinned: false // Will be set by caller using StorageService
      }));

      return {
        worksheets: worksheetList,
        activeSheet: activeWorksheet.name,
        workbookName: workbook.name
      };
    });
  }

  /**
   * Kích hoạt worksheet theo tên
   */
  static async activateWorksheet(sheetName: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(sheetName);
      worksheet.activate();
      await context.sync();
    });
  }

  /**
   * Lấy thông tin cell đang được chọn với thông tin vị trí
   */
  static async getSelectedCellInfo(): Promise<{ value: any; address: string; top: number; left: number; height: number; width: number }> {
    return Excel.run(async (context) => {
      const activeCell = context.workbook.getSelectedRange();
      activeCell.load("values,address,top,left,height,width");
      await context.sync();

      return {
        value: activeCell.values[0][0],
        address: activeCell.address.split('!')[1] || activeCell.address, // Remove sheet name prefix if exists
        top: activeCell.top,
        left: activeCell.left,
        height: activeCell.height,
        width: activeCell.width
      };
    });
  }

  /**
   * Tạo hyperlink trong Excel
   */
  static async createHyperlink(cellAddress: string, linkAddress: string, displayText: string): Promise<void> {
    return await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = worksheet.getRange(cellAddress);
      
      cell.hyperlink = {
        address: linkAddress,
        textToDisplay: displayText
      };
      
      await context.sync();
    });
  }

  /**
   * Kiểm tra worksheet có tồn tại không
   */
  static async worksheetExists(sheetName: string): Promise<boolean> {
    try {
      return await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        
        await context.sync();
        
        return worksheets.items.some(ws => ws.name === sheetName);
      });
    } catch (error) {
      return false;
    }
  }

  /**
   * Thiết lập format cho range
   */
  static async formatRange(rangeName: string, options: {
    fontName?: string;
    fontSize?: number;
    columnWidth?: number;
    rowHeight?: number;
  }): Promise<void> {
    return Excel.run(async (context) => {
      const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = activeWorksheet.getRange(rangeName);
      
      if (options.fontName) {
        range.format.font.name = options.fontName;
      }
      if (options.fontSize) {
        range.format.font.size = options.fontSize;
      }
      if (options.columnWidth) {
        range.format.columnWidth = options.columnWidth;
      }
      if (options.rowHeight) {
        range.format.rowHeight = options.rowHeight;
      }
      
      await context.sync();
    });
  }

  /**
   * Thiết lập page layout
   */
  static async setPageLayout(options: {
    orientation?: Excel.PageOrientation;
    printArea?: string;
  }): Promise<void> {
    return Excel.run(async (context) => {
      const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
      
      try {
        if (options.orientation) {
          activeWorksheet.pageLayout.orientation = options.orientation;
        }
        
        if (options.printArea) {
          const printRange = activeWorksheet.getRange(options.printArea);
          activeWorksheet.pageLayout.setPrintArea(printRange);
        }
        
        await context.sync();
      } catch (error) {
        console.warn("Page layout settings not fully supported:", error);
      }
    });
  }
}

export default ExcelUtils;
