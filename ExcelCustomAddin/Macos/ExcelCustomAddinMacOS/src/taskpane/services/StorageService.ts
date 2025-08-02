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
  private static readonly IMPORTED_FILES_KEY = "ExcelAddin_ImportedFiles";
  private static readonly ALLOW_REIMPORT_KEY = "ExcelAddin_AllowReimport";

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

  /**
   * Lưu thông tin file đã được import
   */
  static addImportedFile(workbookName: string, fileName: string, fileSize: number): void {
    try {
      const key = `${this.IMPORTED_FILES_KEY}_${workbookName}`;
      const data = localStorage.getItem(key);
      const importedFiles = data ? JSON.parse(data) : {};
      
      // Tạo key unique dựa trên tên file và size
      const fileKey = `${fileName}_${fileSize}`;
      importedFiles[fileKey] = {
        fileName,
        fileSize,
        importedAt: new Date().toISOString()
      };
      
      localStorage.setItem(key, JSON.stringify(importedFiles));
    } catch (error) {
      console.error("Error adding imported file:", error);
    }
  }

  /**
   * Kiểm tra file đã được import chưa
   */
  static isFileImported(workbookName: string, fileName: string, fileSize: number): boolean {
    try {
      const key = `${this.IMPORTED_FILES_KEY}_${workbookName}`;
      const data = localStorage.getItem(key);
      if (!data) return false;
      
      const importedFiles = JSON.parse(data);
      const fileKey = `${fileName}_${fileSize}`;
      return fileKey in importedFiles;
    } catch (error) {
      console.error("Error checking imported file:", error);
      return false;
    }
  }

  /**
   * Xóa danh sách file đã import của workbook
   */
  static clearImportedFiles(workbookName: string): void {
    try {
      const key = `${this.IMPORTED_FILES_KEY}_${workbookName}`;
      localStorage.removeItem(key);
    } catch (error) {
      console.error("Error clearing imported files:", error);
    }
  }

  /**
   * Lưu setting cho phép import lại file
   */
  static setAllowReimport(allow: boolean): void {
    try {
      localStorage.setItem(this.ALLOW_REIMPORT_KEY, allow.toString());
    } catch (error) {
      console.error("Error saving allow reimport setting:", error);
    }
  }

  /**
   * Lấy setting cho phép import lại file
   */
  static getAllowReimport(): boolean {
    try {
      const value = localStorage.getItem(this.ALLOW_REIMPORT_KEY);
      return value ? value === 'true' : false;
    } catch (error) {
      console.error("Error getting allow reimport setting:", error);
      return false;
    }
  }
}

export default StorageService;
