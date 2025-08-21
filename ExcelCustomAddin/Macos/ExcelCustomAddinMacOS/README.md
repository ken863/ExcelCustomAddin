# Excel Custom Add-in MacOS - Hướng dẫn Triển khai Docker

## Tổng quan

Dự án Excel Custom Add-in cho MacOS với các tính năng quản lý worksheet và công cụ xử lý hình ảnh nâng cao. Dự án được đóng gói trong Docker để dễ dàng triển khai và phát triển.

## 1. Build và Publish Docker Image

Sử dụng script `build-and-publish.sh` để build và đẩy image lên Docker Hub:

```bash
# Di chuyển đến thư mục chứa script
cd /path/to/ExcelCustomAddin/ExcelCustomAddin/Macos/ExcelCustomAddinMacOS/

# Chạy script build và publish (cần Docker đã login)
DOCKER_USERNAME=your-dockerhub-username \
VERSION=latest \
./build-and-publish.sh
```

### Tham số cấu hình:
- `DOCKER_USERNAME`: Username Docker Hub của bạn (bắt buộc)
- `VERSION`: Tag cho image (mặc định: latest)

## 2. Chạy Container từ Image đã Build

### Chạy container cơ bản:
```bash
docker run -d \
  -p 3000:3000 \
  --name excel-addin \
  --restart unless-stopped \
  your-dockerhub-username/excel-custom-addin:latest
```

### Chạy container với proxy:
```bash
docker run -d \
  -p 3000:3000 \
  --name excel-addin \
  --restart unless-stopped \
  -e HTTP_PROXY=http://user:pass@proxy.company.com:8080 \
  -e HTTPS_PROXY=http://user:pass@proxy.company.com:8080 \
  your-dockerhub-username/excel-custom-addin:latest
```

### Chạy container với certificates nội bộ (mount từ host):
```bash
# Mount certificate từ host và tự động update
docker run -d \
  -p 3000:3000 \
  --name excel-addin \
  --restart unless-stopped \
  -v /path/to/company-cert.crt:/usr/local/share/ca-certificates/company-cert.crt:ro \
  -e HTTP_PROXY=http://user:pass@proxy.company.com:8080 \
  -e HTTPS_PROXY=http://user:pass@proxy.company.com:8080 \
  your-dockerhub-username/excel-custom-addin:latest

# Update certificates sau khi container đã start
docker exec excel-addin update-ca-certificates
```

### Chạy container với certificates sử dụng init script:
```bash
# Mount certificate và sử dụng init script để tự động update
docker run -d \
  -p 3000:3000 \
  --name excel-addin \
  --restart unless-stopped \
  -v /path/to/proxy-cert.crt:/tmp/proxy-cert.crt:ro \
  -e HTTP_PROXY=http://user:pass@proxy.company.com:8080 \
  -e HTTPS_PROXY=http://user:pass@proxy.company.com:8080 \
  -e PROXY_CERT_PATH=/tmp/proxy-cert.crt \
  your-dockerhub-username/excel-custom-addin:latest
```

## 3. Truy cập Ứng dụng

Sau khi container đã khởi động thành công:
- **Ứng dụng**: https://localhost:3000
- **Manifest file**: https://localhost:3000/manifest.xml

## 4. Cài đặt Excel Add-in

### 4.1. Cài đặt qua Manifest URL (Khuyến nghị)

1. **Mở Excel** (Excel for Mac/Web)

2. **Truy cập Insert Tab**:
   - Chọn tab "Insert" trên ribbon
   - Chọn "Get Add-ins" hoặc "Office Add-ins"

3. **Thêm Add-in từ URL**:
   - Chọn "Upload My Add-in"
   - Chọn "Browse..." hoặc "From URL"
   - Nhập URL: `https://localhost:3000/manifest.xml`
   - Chấp nhận cảnh báo certificate (nếu có)

4. **Xác nhận cài đặt**:
   - Add-in sẽ xuất hiện trong tab "Home" ribbon
   - Biểu tượng "Excel Custom Add-in" sẽ hiển thị

### 4.2. Cài đặt thủ công qua File

1. **Tải manifest file**:
   ```bash
   # Tải manifest.xml về máy local
   curl -k https://localhost:3000/manifest.xml -o manifest.xml
   ```

2. **Upload file trong Excel**:
   - Mở Excel
   - Insert → Get Add-ins → Upload My Add-in
   - Chọn "Browse..." và chọn file `manifest.xml` đã tải

### 4.3. Cài đặt cho Development (Excel Desktop)

1. **Thêm shared folder** (Excel for Mac):
   ```bash
   # Tạo thư mục shared manifests
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   
   # Copy manifest file
   cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```

2. **Restart Excel**:
   - Đóng hoàn toàn Excel
   - Mở lại Excel
   - Add-in sẽ tự động được load từ thư mục wef

**Lưu ý cho Excel for Mac**: 
- Excel for Mac không có Trust Center như Windows
- Add-in được tự động trust khi đặt trong thư mục wef
- Nếu vẫn không hiển thị, thử sideload qua Developer tab (nếu có)

### 4.3.1. Phương pháp thay thế cho Excel for Mac

Nếu phương pháp shared folder không hoạt động:

1. **Sử dụng Script Editor** (Office Script):
   ```bash
   # Tạo AppleScript để load add-in
   osascript -e 'tell application "Microsoft Excel"
       activate
       open "https://localhost:3000/manifest.xml"
   end tell'
   ```

2. **Hoặc sử dụng Excel Office Scripts** (nếu có):
   - Mở Excel
   - Automate tab → New Script
   - Paste và chạy script tự động load add-in

### 4.4. Lưu ý đặc biệt cho Excel for Mac

#### Excel for Mac 2019/2021:
- **Không có Trust Center**: Khác với Windows, Excel for Mac không có Trust Center Settings
- **Shared folder path**: Sử dụng `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
- **Auto-trust**: Add-in trong thư mục wef được tự động trust
- **Restart required**: Phải restart Excel sau khi copy manifest vào wef folder

#### Các phương pháp cài đặt theo thứ tự ưu tiên:
1. **URL-based sideloading** (4.1) - Hoạt động tốt nhất
2. **File upload** (4.2) - Backup method 
3. **Shared folder** (4.3) - Cho development

#### Kiểm tra Add-in đã load:
```bash
# Kiểm tra manifest có trong wef folder
ls -la ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/

# Kiểm tra Excel process
ps aux | grep Excel
```

### 4.5. Xử lý lỗi thường gặp

#### Lỗi Certificate/SSL:
- **Nguyên nhân**: Certificate tự ký từ webpack dev server
- **Giải pháp**: 
  1. Truy cập https://localhost:3000 trước trong browser
  2. Chấp nhận certificate warning
  3. Thử lại cài đặt add-in

#### Add-in không hiển thị:
- **Kiểm tra**: Container có đang chạy không (`docker ps`)
- **Kiểm tra**: Port 3000 có accessible không
- **Giải pháp**: Restart container và thử lại

#### Manifest URL không accessible:
```bash
# Test connectivity
curl -k https://localhost:3000/manifest.xml

# Check container logs
docker logs excel-addin
```

### 4.6. Gỡ cài đặt Add-in

1. **Từ Excel**:
   - Insert → My Add-ins
   - Tìm "Excel Custom Add-in"
   - Chọn "..." → Remove

2. **Xóa khỏi shared folder** (nếu cài development):
   ```bash
   rm ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/manifest.xml
   ```

## 5. Quản lý Container

```bash
# Kiểm tra logs
docker logs excel-addin

# Restart container
docker restart excel-addin

# Stop container
docker stop excel-addin

# Remove container
docker rm excel-addin

# Vào trong container để debug
docker exec -it excel-addin /bin/sh
```

## 6. Lưu ý

- Thay đổi `your-dockerhub-username` thành username Docker Hub thực tế của bạn
- Thay đổi thông tin proxy phù hợp với môi trường công ty
- Container sử dụng webpack default certificate để tránh lỗi certificate mismatch
- Đảm bảo port 3000 không bị sử dụng bởi ứng dụng khác
- Container được cấu hình `--restart unless-stopped` để tự động khởi động khi Docker daemon restart
- **Certificates mount**: Có 2 cách mount certificate:
  1. **Direct mount**: Mount trực tiếp vào `/usr/local/share/ca-certificates/` và chạy `update-ca-certificates` manual
  2. **Environment-based**: Mount vào `/tmp/` và set `PROXY_CERT_PATH`, container sẽ tự động install

## 7. Sử dụng Docker Compose

### Development mode:
```bash
# Chạy với docker-compose
cd Docker/
docker-compose up -d

# Để sử dụng proxy certificate, uncomment và edit docker-compose.yml:
# volumes:
#   - ./proxy-cert.crt:/tmp/proxy-cert.crt:ro
# environment:
#   - PROXY_CERT_PATH=/tmp/proxy-cert.crt
```

### Production mode:
```bash
# Sử dụng docker-compose.prod.yml
cd Docker/
docker-compose -f docker-compose.prod.yml up -d

# Set environment variables:
export DOCKER_USERNAME=your-dockerhub-username
export VERSION=latest
export HTTP_PROXY=http://user:pass@proxy.company.com:8080
export HTTPS_PROXY=http://user:pass@proxy.company.com:8080
```
