# Makefile for Excel Custom Add-in

.PHONY: help build start stop restart logs clean install proxy-setup docker-build docker-publish docker-run

# Default target
help:
	@echo "Excel Custom Add-in Commands:"
	@echo "  make install         - Cài đặt dependencies"
	@echo "  make build          - Build Docker image"
	@echo "  make start          - Khởi động add-in server"
	@echo "  make stop           - Dừng add-in server"
	@echo "  make restart        - Restart add-in server"
	@echo "  make logs           - Xem logs real-time"
	@echo "  make clean          - Dọn dẹp containers và volumes"
	@echo "  make dev            - Chạy ở development mode"
	@echo "  make status         - Kiểm tra trạng thái containers"
	@echo "  make proxy-setup    - Thiết lập cấu hình proxy"
	@echo "  make docker-build   - Build Docker image cho production"
	@echo "  make docker-publish - Build và publish lên Docker Hub"
	@echo "  make docker-run     - Chạy từ published Docker image"

# Cài đặt dependencies
install:
	@echo "📦 Installing dependencies..."
	npm install

# Build Docker image
build:
	@echo "🔨 Building Docker image..."
	docker-compose build

# Khởi động services
start:
	@echo "🚀 Starting Excel Add-in server..."
	docker-compose up -d
	@echo "✅ Server started at https://localhost:3000"

# Dừng services
stop:
	@echo "⏹️  Stopping Excel Add-in server..."
	docker-compose down

# Restart services
restart:
	@echo "🔄 Restarting Excel Add-in server..."
	docker-compose restart
	@echo "✅ Server restarted"

# Xem logs
logs:
	@echo "📋 Showing logs (Ctrl+C to exit)..."
	docker-compose logs -f excel-addin

# Development mode (chạy trực tiếp)
dev:
	@echo "🛠️  Starting development server..."
	npm run dev-server

# Kiểm tra status
status:
	@echo "📊 Container status:"
	docker-compose ps

# Dọn dẹp
clean:
	@echo "🧹 Cleaning up containers and volumes..."
	docker-compose down -v
	docker system prune -f

# Setup từ đầu
setup: build start
	@echo "🎉 Setup completed! Add-in is ready at https://localhost:3000"

# Thiết lập proxy
proxy-setup:
	@echo "🔧 Setting up proxy configuration..."
	@if [ ! -f .env ]; then \
		echo "📋 Creating .env file from template..."; \
		cp .env.example .env; \
		echo "✅ Created .env file. Please edit it with your proxy settings:"; \
		echo "   - HTTP_PROXY=http://proxy-server:port"; \
		echo "   - HTTPS_PROXY=http://proxy-server:port"; \
		echo "   - NO_PROXY=localhost,127.0.0.1"; \
		echo ""; \
		echo "📝 Edit .env file: nano .env"; \
	else \
		echo "✅ .env file already exists"; \
		echo "📝 Current proxy settings:"; \
		grep -E "^(HTTP_PROXY|HTTPS_PROXY|NO_PROXY)" .env || echo "   No proxy configured"; \
	fi

# Build Docker image cho production
docker-build:
	@echo "🔨 Building production Docker image..."
	@export DOCKER_USERNAME=$${DOCKER_USERNAME:-your-username} && \
	docker build \
		--build-arg BUILD_DATE="$(shell date -u +'%Y-%m-%dT%H:%M:%SZ')" \
		--build-arg VCS_REF="$(shell git rev-parse --short HEAD 2>/dev/null || echo 'unknown')" \
		--tag "$$DOCKER_USERNAME/excel-custom-addin:latest" \
		--tag "$$DOCKER_USERNAME/excel-custom-addin:$(shell date +%Y%m%d)" \
		.
	@echo "✅ Build completed!"

# Build và publish lên Docker Hub
docker-publish:
	@echo "🚀 Building and publishing to Docker Hub..."
	@if [ -z "$$DOCKER_USERNAME" ]; then \
		echo "❌ DOCKER_USERNAME environment variable is required"; \
		echo "   Set it with: export DOCKER_USERNAME=your-dockerhub-username"; \
		exit 1; \
	fi
	./build-and-publish.sh

# Chạy từ published Docker image
docker-run:
	@echo "🚀 Running from published Docker image..."
	@export DOCKER_USERNAME=$${DOCKER_USERNAME:-your-username} && \
	docker-compose -f docker-compose.prod.yml up -d
	@echo "✅ Started production container from Docker Hub image"
