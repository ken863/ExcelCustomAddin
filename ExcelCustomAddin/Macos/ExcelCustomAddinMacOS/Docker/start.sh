#!/bin/sh

# Script khởi động cho Excel Custom Add-in
# Sử dụng webpack default certificate generation để tránh complexity

echo "🚀 Starting Excel Custom Add-in..."

# Xử lý certificate từ mount path nếu có
if [ -n "$PROXY_CERT_PATH" ] && [ -f "$PROXY_CERT_PATH" ]; then
    echo "📜 Installing proxy certificate from $PROXY_CERT_PATH..."
    cp "$PROXY_CERT_PATH" /usr/local/share/ca-certificates/
    update-ca-certificates
    echo "✅ Proxy certificate installed successfully"
fi



# Thiết lập npm config cho user
npm config set cache /home/nodeuser/.npm
npm config set prefix /home/nodeuser/.npm-global

# Sử dụng webpack config đơn giản cho Docker để tránh certificate issues
if [ -f "/app/webpack.config.simple.js" ]; then
    echo "🐳 Using simplified Docker webpack configuration (avoids certificate issues)"
    export WEBPACK_CONFIG_PATH="/app/webpack.config.simple.js"
    # Backup original và sử dụng simple config
    if [ ! -f "/app/webpack.config.original.js" ]; then
        mv /app/webpack.config.js /app/webpack.config.original.js
    fi
    ln -sf /app/webpack.config.simple.js /app/webpack.config.js
    
    echo "✅ Webpack will generate its own self-signed certificate"
else
    echo "📝 Using original webpack configuration"
fi

echo "🌐 Starting development server..."

# Khởi động ứng dụng
exec "$@"
