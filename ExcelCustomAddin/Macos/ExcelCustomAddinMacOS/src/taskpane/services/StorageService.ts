/**
 * Service để quản lý việc lưu trữ và truy xuất thông tin pin sheets
 * Tương đương với Dictionary<string, HashSet<string>> PinnedSheets trong C#
 */

interface PinnedSheetsData {
  [workbookName: string]: string[];
}

class StorageService {
  private static readonly PINNED_SHEETS_KEY = "ExcelAddin_PinnedSheets";
  private static readonly SCALE_PERCENT_KEY = "ExcelAddin_ScalePercent";

  /**
   * Lấy danh sách sheets được pin cho workbook
   */
  static getPinnedSheets(workbookName: string): Set<string> {
    try {
      const data = localStorage.getItem(this.PINNED_SHEETS_KEY);
      if (!data) return new Set();

      const pinnedData: PinnedSheetsData = JSON.parse(data);
      const sheets = pinnedData[workbookName] || [];
      return new Set(sheets);
    } catch (error) {
      console.error("Error getting pinned sheets:", error);
      return new Set();
    }
  }

  /**
   * Toggle pin status của sheet
   */
  static togglePinSheet(workbookName: string, sheetName: string): boolean {
    try {
      const data = localStorage.getItem(this.PINNED_SHEETS_KEY);
      const pinnedData: PinnedSheetsData = data ? JSON.parse(data) : {};

      if (!pinnedData[workbookName]) {
        pinnedData[workbookName] = [];
      }

      const sheets = new Set(pinnedData[workbookName]);
      let isPinned = false;

      if (sheets.has(sheetName)) {
        sheets.delete(sheetName);
        isPinned = false;
      } else {
        sheets.add(sheetName);
        isPinned = true;
      }

      pinnedData[workbookName] = Array.from(sheets);
      localStorage.setItem(this.PINNED_SHEETS_KEY, JSON.stringify(pinnedData));

      return isPinned;
    } catch (error) {
      console.error("Error toggling pin sheet:", error);
      return false;
    }
  }

  /**
   * Kiểm tra sheet có được pin không
   */
  static isSheetPinned(workbookName: string, sheetName: string): boolean {
    const pinnedSheets = this.getPinnedSheets(workbookName);
    return pinnedSheets.has(sheetName);
  }

  /**
   * Xóa tất cả pinned sheets của workbook (khi workbook đóng)
   */
  static clearPinnedSheets(workbookName: string): void {
    try {
      const data = localStorage.getItem(this.PINNED_SHEETS_KEY);
      if (!data) return;

      const pinnedData: PinnedSheetsData = JSON.parse(data);
      delete pinnedData[workbookName];
      localStorage.setItem(this.PINNED_SHEETS_KEY, JSON.stringify(pinnedData));
    } catch (error) {
      console.error("Error clearing pinned sheets:", error);
    }
  }

  /**
   * Lưu scale percent
   */
  static saveScalePercent(percent: number): void {
    try {
      localStorage.setItem(this.SCALE_PERCENT_KEY, percent.toString());
    } catch (error) {
      console.error("Error saving scale percent:", error);
    }
  }

  /**
   * Lấy scale percent
   */
  static getScalePercent(): number {
    try {
      const value = localStorage.getItem(this.SCALE_PERCENT_KEY);
      return value ? parseInt(value, 10) : 85;
    } catch (error) {
      console.error("Error getting scale percent:", error);
      return 85;
    }
  }
}

export default StorageService;
