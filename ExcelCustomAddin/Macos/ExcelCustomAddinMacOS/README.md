# Hướng dẫn Build & Publish Docker Image

## 1. Build và Publish image lên Docker Hub

```bash
# Chạy script build và publish (cần Docker đã login)
DOCKER_USERNAME=your-dockerhub-username VERSION=latest \
HTTP_PROXY=http://user:pass@proxy.company.com:8080 \
HTTPS_PROXY=http://user:pass@proxy.company.com:8080 \
./build-and-publish.sh
```

- Tham số `DOCKER_USERNAME` là username Docker Hub của bạn.
- Tham số `VERSION` là tag cho image (mặc định: latest).
- Có thể truyền thêm biến môi trường proxy nếu cần.

## 2. Tạo container từ image đã publish với HTTP proxy

```bash
# Kéo image từ Docker Hub
# docker pull your-dockerhub-username/excel-custom-addin:latest

# Chạy container với proxy

docker run -d \
  -p 3000:3000 \
  -e HTTP_PROXY=http://user:pass@proxy.company.com:8080 \
  -e HTTPS_PROXY=http://user:pass@proxy.company.com:8080 \
  --name excel-addin \
  your-dockerhub-username/excel-custom-addin:latest
```

- Thay đổi thông tin proxy, username, password cho phù hợp môi trường của bạn.
- Nếu cần thêm certificate nội bộ, copy file `.crt` vào thư mục Docker trước khi build (ví dụ: Docker/your-cert.crt).
