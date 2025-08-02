/**
 * Event handlers for Excel operations
 */

import StorageService from "../services/StorageService";

export class EventHandlers {
  private static sheetActivatedHandlers: Set<() => void> = new Set();
  private static sheetAddedHandlers: Set<() => void> = new Set();
  private static sheetDeletedHandlers: Set<() => void> = new Set();
  private static workbookClosedHandlers: Set<(workbookName: string) => void> = new Set();
  private static intervalId: NodeJS.Timeout | null = null;
  private static currentWorkbookName: string = "";

  /**
   * Đăng ký event listener cho worksheet activated
   */
  static registerSheetActivatedHandler(handler: () => void): () => void {
    this.sheetActivatedHandlers.add(handler);
    
    // Return unregister function
    return () => {
      this.sheetActivatedHandlers.delete(handler);
    };
  }

  /**
   * Đăng ký event listener cho worksheet added
   */
  static registerSheetAddedHandler(handler: () => void): () => void {
    this.sheetAddedHandlers.add(handler);
    
    // Return unregister function
    return () => {
      this.sheetAddedHandlers.delete(handler);
    };
  }

  /**
   * Đăng ký event listener cho worksheet deleted
   */
  static registerSheetDeletedHandler(handler: () => void): () => void {
    this.sheetDeletedHandlers.add(handler);
    
    // Return unregister function
    return () => {
      this.sheetDeletedHandlers.delete(handler);
    };
  }

  /**
   * Đăng ký event listener cho workbook closed
   */
  static registerWorkbookClosedHandler(handler: (workbookName: string) => void): () => void {
    this.workbookClosedHandlers.add(handler);
    
    // Return unregister function
    return () => {
      this.workbookClosedHandlers.delete(handler);
    };
  }

  /**
   * Khởi tạo workbook monitoring để detect khi workbook đóng
   */
  static initializeWorkbookMonitoring(): void {
    // Monitor cho beforeunload event (khi tab/window đóng)
    window.addEventListener('beforeunload', this.handleWorkbookClose.bind(this));
    
    // Monitor cho visibility change (khi switch tab)
    document.addEventListener('visibilitychange', this.handleVisibilityChange.bind(this));
    
    // Lưu workbook name hiện tại
    this.updateCurrentWorkbookName();
    
    console.log("Workbook monitoring initialized");
  }

  /**
   * Cập nhật tên workbook hiện tại
   */
  private static async updateCurrentWorkbookName(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        workbook.load("name");
        await context.sync();
        this.currentWorkbookName = workbook.name;
      });
    } catch (error) {
      console.error("Error getting current workbook name:", error);
    }
  }

  /**
   * Xử lý khi workbook có thể đóng
   */
  private static handleWorkbookClose(): void {
    if (this.currentWorkbookName) {
      console.log(`Workbook "${this.currentWorkbookName}" is closing, clearing pinned sheets storage`);
      
      // Clear storage cho workbook hiện tại
      StorageService.clearPinnedSheets(this.currentWorkbookName);
      
      // Trigger handlers
      this.workbookClosedHandlers.forEach(handler => {
        try {
          handler(this.currentWorkbookName);
        } catch (error) {
          console.error("Error in workbook closed handler:", error);
        }
      });
    }
  }

  /**
   * Xử lý khi visibility thay đổi (có thể là workbook switch)
   */
  private static handleVisibilityChange(): void {
    if (document.visibilityState === 'hidden') {
      // Document hidden, có thể workbook đang đóng hoặc switch
      // Delay một chút để check nếu thực sự đóng
      setTimeout(() => {
        if (document.visibilityState === 'hidden') {
          this.handleWorkbookClose();
        }
      }, 1000);
    } else if (document.visibilityState === 'visible') {
      // Document visible lại, cập nhật workbook name
      this.updateCurrentWorkbookName();
    }
  }

  /**
   * Khởi tạo Excel event listeners
   */
  static async initializeExcelEvents(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        // Register worksheet events
        const worksheets = context.workbook.worksheets;
        
        // Sheet activated event
        worksheets.onActivated.add(async () => {
          this.sheetActivatedHandlers.forEach(handler => {
            try {
              handler();
            } catch (error) {
              console.error("Error in sheet activated handler:", error);
            }
          });
        });

        // Sheet added event
        worksheets.onAdded.add(async () => {
          this.sheetAddedHandlers.forEach(handler => {
            try {
              handler();
            } catch (error) {
              console.error("Error in sheet added handler:", error);
            }
          });
        });

        // Sheet deleted event
        worksheets.onDeleted.add(async () => {
          this.sheetDeletedHandlers.forEach(handler => {
            try {
              handler();
            } catch (error) {
              console.error("Error in sheet deleted handler:", error);
            }
          });
        });

        await context.sync();
        console.log("Excel events initialized successfully");
      });
    } catch (error) {
      console.error("Error initializing Excel events:", error);
      // Fallback to polling if events are not available
      this.startPollingFallback();
    }
  }

  /**
   * Bắt đầu auto-refresh với polling (fallback method)
   */
  static startPollingFallback(intervalMs: number = 5000): void {
    if (this.intervalId) {
      clearInterval(this.intervalId);
    }

    this.intervalId = setInterval(() => {
      // Trigger all handlers as a general refresh
      this.sheetActivatedHandlers.forEach(handler => {
        try {
          handler();
        } catch (error) {
          console.error("Error in polling handler:", error);
        }
      });
    }, intervalMs);

    console.log(`Started polling fallback with ${intervalMs}ms interval`);
  }

  /**
   * Dừng auto-refresh polling
   */
  static stopPolling(): void {
    if (this.intervalId) {
      clearInterval(this.intervalId);
      this.intervalId = null;
      console.log("Stopped polling");
    }
  }

  /**
   * Làm sạch tất cả event handlers
   */
  static cleanup(): void {
    this.sheetActivatedHandlers.clear();
    this.sheetAddedHandlers.clear();
    this.sheetDeletedHandlers.clear();
    this.workbookClosedHandlers.clear();
    this.stopPolling();
    
    // Remove event listeners
    window.removeEventListener('beforeunload', this.handleWorkbookClose.bind(this));
    document.removeEventListener('visibilitychange', this.handleVisibilityChange.bind(this));
    
    console.log("Event handlers cleaned up");
  }

  /**
   * Kiểm tra trạng thái của event handlers
   */
  static getStatus(): {
    activatedHandlers: number;
    addedHandlers: number;
    deletedHandlers: number;
    workbookClosedHandlers: number;
    isPolling: boolean;
    currentWorkbook: string;
  } {
    return {
      activatedHandlers: this.sheetActivatedHandlers.size,
      addedHandlers: this.sheetAddedHandlers.size,
      deletedHandlers: this.sheetDeletedHandlers.size,
      workbookClosedHandlers: this.workbookClosedHandlers.size,
      isPolling: this.intervalId !== null,
      currentWorkbook: this.currentWorkbookName
    };
  }

  /**
   * Xử lý Excel events với error handling
   */
  static async safeExcelRun<T>(operation: (context: Excel.RequestContext) => Promise<T>): Promise<T | null> {
    try {
      return await Excel.run(operation);
    } catch (error) {
      console.error("Excel operation failed:", error);
      return null;
    }
  }

  /**
   * Thực hiện batch operations với error handling
   */
  static async batchOperations(operations: Array<() => Promise<any>>): Promise<void> {
    const results = await Promise.allSettled(operations.map(op => op()));
    
    const failures = results
      .map((result, index) => ({ result, index }))
      .filter(({ result }) => result.status === 'rejected');

    if (failures.length > 0) {
      console.warn(`${failures.length} operations failed:`, failures);
    }
  }

  /**
   * Retry operation với exponential backoff
   */
  static async retryOperation<T>(
    operation: () => Promise<T>,
    maxRetries: number = 3,
    baseDelay: number = 1000
  ): Promise<T | null> {
    for (let attempt = 0; attempt < maxRetries; attempt++) {
      try {
        return await operation();
      } catch (error) {
        if (attempt === maxRetries - 1) {
          console.error(`Operation failed after ${maxRetries} attempts:`, error);
          return null;
        }
        
        const delay = baseDelay * Math.pow(2, attempt);
        console.warn(`Operation failed (attempt ${attempt + 1}), retrying in ${delay}ms:`, error);
        await new Promise(resolve => setTimeout(resolve, delay));
      }
    }
    return null;
  }
}

export default EventHandlers;
