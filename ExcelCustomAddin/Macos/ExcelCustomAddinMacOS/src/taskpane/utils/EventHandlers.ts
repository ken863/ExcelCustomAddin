/**
 * Event handlers for Excel operations
 */

export class EventHandlers {
  private static sheetActivatedHandlers: Set<() => void> = new Set();
  private static sheetAddedHandlers: Set<() => void> = new Set();
  private static sheetDeletedHandlers: Set<() => void> = new Set();
  private static intervalId: NodeJS.Timeout | null = null;

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
    this.stopPolling();
    console.log("Event handlers cleaned up");
  }

  /**
   * Kiểm tra trạng thái của event handlers
   */
  static getStatus(): {
    activatedHandlers: number;
    addedHandlers: number;
    deletedHandlers: number;
    isPolling: boolean;
  } {
    return {
      activatedHandlers: this.sheetActivatedHandlers.size,
      addedHandlers: this.sheetAddedHandlers.size,
      deletedHandlers: this.sheetDeletedHandlers.size,
      isPolling: this.intervalId !== null
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
