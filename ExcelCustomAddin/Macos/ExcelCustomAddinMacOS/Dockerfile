## moved to Docker/Dockerfile
FROM node:18-alpine

# Metadata
ARG BUILD_DATE
ARG VCS_REF
ARG VERSION=1.0.0

LABEL maintainer="Excel Custom Add-in Team" \
      org.opencontainers.image.title="Excel Custom Add-in" \
      org.opencontainers.image.description="Advanced worksheet management and image tools for Excel workbooks" \
      org.opencontainers.image.version=${VERSION} \
      org.opencontainers.image.created=${BUILD_DATE} \
      org.opencontainers.image.revision=${VCS_REF} \
      org.opencontainers.image.url="https://github.com/ken863/ExcelCustomAddin" \
      org.opencontainers.image.source="https://github.com/ken863/ExcelCustomAddin" \
      org.opencontainers.image.vendor="Excel Custom Add-in Team" \
      org.opencontainers.image.licenses="MIT"

# Thiết lập proxy arguments
ARG HTTP_PROXY
ARG HTTPS_PROXY
ARG http_proxy
ARG https_proxy

# Set proxy environment variables
ENV HTTP_PROXY=${HTTP_PROXY}
ENV HTTPS_PROXY=${HTTPS_PROXY}
ENV http_proxy=${http_proxy}
ENV https_proxy=${https_proxy}

# Thiết lập thư mục làm việc
WORKDIR /app

# Tạo user non-root cho security và cấu hình sudo
RUN addgroup -g 1001 -S nodejs && \
    adduser -S nodeuser -u 1001 && \
    echo 'nodeuser ALL=(ALL) NOPASSWD: ALL' >> /etc/sudoers

# Cài đặt các dependencies hệ thống cần thiết
RUN apk add --no-cache \
    openssl \
    ca-certificates \
    dumb-init \
    wget \
    git \
    python3 \
    make \
    g++ \
    sudo

# Copy toàn bộ context để có thể copy certificates nếu có
COPY . /tmp/build-context/

# Copy certificates nếu có và cấu hình npm cho proxy environment
RUN if ls /tmp/build-context/*.crt 1> /dev/null 2>&1; then \
        echo "📜 Found certificate files, copying..."; \
        cp /tmp/build-context/*.crt /usr/local/share/ca-certificates/; \
        update-ca-certificates; \
    else \
        echo "📜 No certificate files found"; \
    fi && \
    if [ -n "$HTTP_PROXY" ] || [ -n "$HTTPS_PROXY" ]; then \
        echo "🔧 Configuring npm for proxy environment..."; \
        npm config set strict-ssl false; \
        npm config set registry https://registry.npmjs.org/; \
    fi

# Copy package files và cài đặt dependencies
RUN cp /tmp/build-context/package*.json ./ && \
    npm install --silent

# Copy toàn bộ source code từ temp location
RUN cp -r /tmp/build-context/src ./src && \
    cp -r /tmp/build-context/assets ./assets && \
    cp /tmp/build-context/*.json . && \
    cp /tmp/build-context/webpack.config.js . && \
    cp /tmp/build-context/babel.config.json . && \
    cp /tmp/build-context/manifest.xml . && \
    chown -R nodeuser:nodejs /app && \
    rm -rf /tmp/build-context

# Build ứng dụng
RUN npm run build

# Tạo thư mục cần thiết cho certificates và set permissions
RUN mkdir -p /home/nodeuser/.office-addin-dev-certs && \
    chown -R nodeuser:nodejs /home/nodeuser/.office-addin-dev-certs && \
    chown -R nodeuser:nodejs /app

# Chuyển sang user non-root
USER nodeuser

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=30s --retries=3 \
    CMD wget --no-verbose --tries=1 --spider http://localhost:3000/ || exit 1

# Expose port
EXPOSE 3000

# Sử dụng dumb-init để handle signals properly
ENTRYPOINT ["dumb-init", "--"]

# Start the application
CMD ["npm", "run", "dev-server"]
