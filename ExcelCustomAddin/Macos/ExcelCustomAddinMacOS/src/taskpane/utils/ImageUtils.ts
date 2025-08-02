/**
 * Image utilities for Excel operations
 */

import StorageService from '../services/StorageService';

export class ImageUtils {
  /**
   * Chuyển file thành base64 string
   */
  static fileToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        resolve(result);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  /**
   * Import tất cả hình ảnh từ file picker
   */
  static async importImagesFromFolder(): Promise<{ success: boolean; message: string }> {
    try {
      // Sử dụng file picker trực tiếp
      return this.openFilePicker();
      
    } catch (error) {
      console.error("Error importing images:", error);
      return {
        success: false,
        message: `Có lỗi xảy ra khi import hình ảnh: ${error.message}`
      };
    }
  }

  /**
   * Mở file picker với gợi ý folder (nếu có)
   */
  static async openFilePickerWithHint(suggestedPath?: string): Promise<{ success: boolean; message: string }> {
    return new Promise((resolve) => {
      try {
        // Hiển thị thông báo nếu có suggested path
        if (suggestedPath && suggestedPath.trim() !== '') {
          const userWantsToNavigate = confirm(
            `Đường dẫn được đề xuất: ${suggestedPath}\n\n` +
            `Nhấn OK để mở file picker và navigate đến folder này.\n` +
            `Nhấn Cancel để mở file picker ở vị trí mặc định.`
          );
          
          if (!userWantsToNavigate) {
            // User chọn Cancel, chỉ mở file picker bình thường
          }
        }

        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.multiple = true;
        fileInput.accept = '.jpg,.jpeg,.png,.bmp,.gif,.tiff,image/*';
        
        fileInput.onchange = async (event) => {
          const files = (event.target as HTMLInputElement).files;
          if (!files || files.length === 0) {
            resolve({
              success: false,
              message: "Không có file nào được chọn."
            });
            return;
          }

          const result = await this.insertMultipleImages(files);
          resolve(result);
        };

        fileInput.oncancel = () => {
          resolve({
            success: false,
            message: "Đã hủy chọn file."
          });
        };

        fileInput.click();
      } catch (error) {
        resolve({
          success: false,
          message: `Có lỗi xảy ra: ${error.message}`
        });
      }
    });
  }

  /**
   * Mở file picker để chọn nhiều hình ảnh (fallback method)
   */
  static async openFilePicker(): Promise<{ success: boolean; message: string }> {
    return new Promise((resolve) => {
      try {
        const fileInput = document.createElement('input');
        fileInput.type = 'file';
        fileInput.multiple = true;
        fileInput.accept = '.jpg,.jpeg,.png,.bmp,.gif,.tiff,image/*';
        
        fileInput.onchange = async (event) => {
          const files = (event.target as HTMLInputElement).files;
          
          if (!files || files.length === 0) {
            resolve({
              success: false,
              message: "Không có file nào được chọn."
            });
            return;
          }

          const result = await this.insertMultipleImages(files);
          resolve(result);
        };

        fileInput.oncancel = () => {
          resolve({
            success: false,
            message: "Đã hủy chọn file."
          });
        };

        fileInput.click();
      } catch (error) {
        resolve({
          success: false,
          message: `Có lỗi xảy ra: ${error.message}`
        });
      }
    });
  }

  /**
   * Helper function để tạo FileList từ array of Files
   */
  static createFileList(files: File[]): FileList {
    const dataTransfer = new DataTransfer();
    files.forEach(file => dataTransfer.items.add(file));
    return dataTransfer.files;
  }
  static async insertMultipleImages(files: FileList): Promise<{ success: boolean; message: string }> {
    try {
      if (!files || files.length === 0) {
        return {
          success: false,
          message: "Không có file nào được chọn."
        };
      }

      // Lọc chỉ lấy file hình ảnh
      const imageFiles = Array.from(files).filter(file => 
        file.type.startsWith('image/')
      );

      if (imageFiles.length === 0) {
        return {
          success: false,
          message: "Không có file hình ảnh hợp lệ nào được chọn."
        };
      }

      const scalePercent = StorageService.getScalePercent();
      const scale = scalePercent / 100;
      const allowReimport = StorageService.getAllowReimport();

      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const workbook = context.workbook;
        worksheet.load("name");
        workbook.load("name");
        
        // Lấy thông tin về ô hiện tại để làm điểm bắt đầu
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("address, left, top, width, height, rowIndex, columnIndex");
        
        await context.sync();

        const workbookName = workbook.name;
        let currentRow = selectedRange.rowIndex;
        let currentCol = selectedRange.columnIndex;
        let currentVerticalOffset = 0; // Track vertical offset from starting position
        let importedCount = 0;
        let skippedCount = 0;
        let reimportedCount = 0;
        const skippedFiles: string[] = [];
        const reimportedFiles: string[] = [];

        // Xử lý từng file
        for (let i = 0; i < imageFiles.length; i++) {
          const file = imageFiles[i];
          
          try {
            // Kiểm tra file đã được import chưa
            const isAlreadyImported = StorageService.isFileImported(workbookName, file.name, file.size);
            
            if (isAlreadyImported && !allowReimport) {
              skippedCount++;
              skippedFiles.push(file.name);
              continue;
            }
            
            // Đánh dấu nếu là file import lại
            if (isAlreadyImported && allowReimport) {
              reimportedCount++;
              reimportedFiles.push(file.name);
            }
            
            // Chuyển file thành base64
            const base64 = await this.fileToBase64(file);
            const base64Data = base64.split(',')[1]; // Remove data:image/...;base64, prefix
            
            // Tính toán vị trí chèn hình
            const targetCell = worksheet.getCell(currentRow, currentCol);
            targetCell.load("left, top, width, height");
            await context.sync();
            
            // Chèn hình ảnh
            const image = worksheet.shapes.addImage(base64Data);
            image.load("id, height, width, left, top");
            await context.sync();
            
            // Scale hình ảnh
            const originalWidth = image.width;
            const originalHeight = image.height;
            const scaledWidth = originalWidth * scale;
            const scaledHeight = originalHeight * scale;
            
            image.width = scaledWidth;
            image.height = scaledHeight;
            
            // Đặt vị trí hình ảnh (vertical arrangement)
            image.left = targetCell.left;
            image.top = targetCell.top + currentVerticalOffset;
            
            // Cập nhật tên hình ảnh
            image.name = `Image_${importedCount + 1}_${file.name}`;
            
            await context.sync();
            
            // Lưu thông tin file đã import
            StorageService.addImportedFile(workbookName, file.name, file.size);
            
            // Di chuyển đến vị trí tiếp theo (xuống dưới)
            currentVerticalOffset += scaledHeight + 10; // Add 10px spacing between images
            importedCount++;
            
          } catch (fileError) {
            console.error(`Error processing file ${file.name}:`, fileError);
            // Tiếp tục với file tiếp theo
          }
        }

        let message = "";
        
        if (importedCount > 0) {
          message = `Đã chèn ${importedCount} hình ảnh thành công với scale ${scalePercent}%!`;
        }
        
        if (reimportedCount > 0) {
          if (message) message += "\n\n";
          message += `${reimportedCount} file đã được import lại:`;
          reimportedFiles.forEach(fileName => {
            message += `\n• ${fileName}`;
          });
        }
        
        if (skippedCount > 0) {
          if (message) message += "\n\n";
          message += `${skippedCount} file đã bị bỏ qua vì đã được import trước đó:`;
          skippedFiles.forEach(fileName => {
            message += `\n• ${fileName}`;
          });
        }
        
        // Nếu không có file nào được import và không có file nào bị skip
        if (importedCount === 0 && skippedCount === 0 && reimportedCount === 0) {
          message = "Không có file nào được xử lý.";
        }

        return {
          success: true,
          message: message
        };
      });
      
    } catch (error) {
      console.error("Error inserting images:", error);
      return {
        success: false,
        message: `Có lỗi xảy ra khi chèn hình ảnh: ${error.message}`
      };
    }
  }

  /**
   * Chèn một hình ảnh vào vị trí hiện tại
   */
  static async insertSingleImage(file: File, customScale?: number): Promise<{ success: boolean; message: string }> {
    try {
      if (!file.type.startsWith('image/')) {
        return {
          success: false,
          message: "File được chọn không phải là hình ảnh hợp lệ."
        };
      }

      const scalePercent = customScale || StorageService.getScalePercent();
      const scale = scalePercent / 100;

      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load("left, top");
        
        await context.sync();

        // Chuyển file thành base64
        const base64 = await this.fileToBase64(file);
        const base64Data = base64.split(',')[1];
        
        // Chèn hình ảnh
        const image = worksheet.shapes.addImage(base64Data);
        image.load("id, height, width");
        await context.sync();
        
        // Scale và đặt vị trí
        image.width = image.width * scale;
        image.height = image.height * scale;
        image.left = selectedRange.left;
        image.top = selectedRange.top;
        image.name = `Image_${file.name}`;
        
        await context.sync();

        return {
          success: true,
          message: `Đã chèn hình ảnh '${file.name}' thành công với scale ${scalePercent}%!`
        };
      });
      
    } catch (error) {
      console.error("Error inserting single image:", error);
      return {
        success: false,
        message: `Có lỗi xảy ra khi chèn hình ảnh: ${error.message}`
      };
    }
  }

  /**
   * Xóa tất cả hình ảnh trong worksheet hiện tại
   */
  static async clearAllImages(): Promise<{ success: boolean; message: string }> {
    try {
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const workbook = context.workbook;
        const shapes = worksheet.shapes;
        shapes.load("items/type");
        workbook.load("name");
        
        await context.sync();
        
        let imageCount = 0;
        
        // Xóa tất cả shapes có type là image
        for (const shape of shapes.items) {
          if (shape.type === "Image") {
            shape.delete();
            imageCount++;
          }
        }
        
        await context.sync();

        // Xóa luôn danh sách file đã import của workbook này
        StorageService.clearImportedFiles(workbook.name);

        return {
          success: true,
          message: `Đã xóa ${imageCount} hình ảnh khỏi worksheet và reset danh sách file đã import.`
        };
      });
      
    } catch (error) {
      console.error("Error clearing images:", error);
      return {
        success: false,
        message: `Có lỗi xảy ra khi xóa hình ảnh: ${error.message}`
      };
    }
  }

  /**
   * Resize tất cả hình ảnh trong worksheet theo scale mới
   */
  static async resizeAllImages(newScalePercent: number): Promise<{ success: boolean; message: string }> {
    try {
      const scale = newScalePercent / 100;
      
      return await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const shapes = worksheet.shapes;
        shapes.load("items/type");
        
        await context.sync();
        
        let imageCount = 0;
        
        // Resize tất cả images
        for (const shape of shapes.items) {
          if (shape.type === "Image") {
            shape.load("width, height");
            await context.sync();
            
            // Giả sử scale hiện tại và tính toán lại
            shape.scaleWidth(scale, Excel.ShapeScaleType.currentSize);
            shape.scaleHeight(scale, Excel.ShapeScaleType.currentSize);
            imageCount++;
          }
        }
        
        await context.sync();

        return {
          success: true,
          message: `Đã thay đổi kích thước ${imageCount} hình ảnh theo scale ${newScalePercent}%.`
        };
      });
      
    } catch (error) {
      console.error("Error resizing images:", error);
      return {
        success: false,
        message: `Có lỗi xảy ra khi thay đổi kích thước hình ảnh: ${error.message}`
      };
    }
  }
}

export default ImageUtils;
