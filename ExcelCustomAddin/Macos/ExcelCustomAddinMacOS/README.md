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

## 4. Quản lý Container

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

## 5. Lưu ý

- Thay đổi `your-dockerhub-username` thành username Docker Hub thực tế của bạn
- Thay đổi thông tin proxy phù hợp với môi trường công ty
- Container sử dụng webpack default certificate để tránh lỗi certificate mismatch
- Đảm bảo port 3000 không bị sử dụng bởi ứng dụng khác
- Container được cấu hình `--restart unless-stopped` để tự động khởi động khi Docker daemon restart
- **Certificates mount**: Có 2 cách mount certificate:
  1. **Direct mount**: Mount trực tiếp vào `/usr/local/share/ca-certificates/` và chạy `update-ca-certificates` manual
  2. **Environment-based**: Mount vào `/tmp/` và set `PROXY_CERT_PATH`, container sẽ tự động install

## 6. Sử dụng Docker Compose

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
