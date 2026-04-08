#!/bin/bash
# =============================================================
# 报销费用填写工具 · 一键安装脚本（Linux/macOS 服务器）
# =============================================================
# 用法: bash install.sh
#
# 适用系统: Ubuntu 20.04+ / Debian 11+ / CentOS 8+ / macOS
# 完成后请阅读「配置指南」章节
# =============================================================

set -e  # 遇错即停

echo "============================================"
echo "报销费用填写工具 · 云端部署脚本"
echo "============================================"
echo ""

# ---- 颜色 ----
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; NC='\033[0m'

info()    { echo -e "${GREEN}[INFO]${NC} $1"; }
warn()    { echo -e "${YELLOW}[WARN]${NC} $1"; }
error()   { echo -e "${RED}[ERROR]${NC} $1"; exit 1; }

# ---- 检测操作系统 ----
detect_os() {
    if [[ "$OSTYPE" == "linux-gnu"* ]]; then
        if [[ -f /etc/debian_version ]]; then
            PKG="apt-get"; DIST="Debian/Ubuntu"
        elif [[ -f /etc/redhat-release ]]; then
            PKG="yum";     DIST="CentOS/RHEL"
        else
            PKG="unknown"; DIST="Linux"
        fi
    elif [[ "$OSTYPE" == "darwin"* ]]; then
        PKG="brew"; DIST="macOS"
    else
        error "不支持的操作系统: $OSTYPE"
    fi
    info "检测到系统: $DIST ($PKG)"
}

# ---- 1. Python 环境检查 ----
check_python() {
    info "检查 Python 环境..."
    if command -v python3 &> /dev/null; then
        PYTHON=$(command -v python3)
    elif command -v python &> /dev/null; then
        PYTHON=$(command -v python)
    else
        error "未找到 Python，请先安装 Python 3.8+"
    fi
    VER=$($PYTHON --version 2>&1 | awk '{print $2}')
    info "Python 版本: $VER"
    if ! $PYTHON -c "import sys; sys.exit(0 if sys.version_info >= (3,8) else 1)"; then
        error "Python 版本过低，需要 Python 3.8 以上"
    fi
}

# ---- 2. 安装系统依赖 ----
install_system_deps() {
    info "安装系统依赖..."
    if [[ "$PKG" == "apt-get" ]]; then
        sudo apt-get update -qq
        sudo apt-get install -y -qq python3-pip python3-venv \
            libgl1-mesa-glx libglib2.0-0 > /dev/null 2>&1
    elif [[ "$PKG" == "yum" ]]; then
        sudo yum install -y -q python3 python3-pip > /dev/null 2>&1
    elif [[ "$PKG" == "brew" ]]; then
        brew install python3 > /dev/null 2>&1 || true
    fi
    info "系统依赖安装完成"
}

# ---- 3. 创建虚拟环境（可选，推荐）----
setup_venv() {
    info "创建 Python 虚拟环境..."
    if [[ -d "venv" ]]; then
        warn "虚拟环境已存在，跳过创建"
    else
        $PYTHON -m venv venv
        info "虚拟环境创建完成: ./venv"
    fi
    PIP="./venv/bin/pip"
    PY="./venv/bin/python"
}

# ---- 4. 安装 Python 依赖 ----
install_python_deps() {
    info "安装 Python 依赖..."
    # 优先用 venv，否则用系统 pip
    if [[ -f "venv/bin/pip" ]]; then
        PIP="$PWD/venv/bin/pip"
        PY="$PWD/venv/bin/python"
    else
        PIP="$PYTHON -m pip"
        PY="$PYTHON"
    fi

    $PIP install --quiet --upgrade pip
    $PIP install --quiet openpyxl>=3.1.0 Pillow>=10.0.0

    # 验证
    if $PY -c "import openpyxl, PIL; print('OK')" > /dev/null 2>&1; then
        info "Python 依赖安装成功 ✅"
    else
        error "依赖安装验证失败，请手动运行: pip install openpyxl Pillow"
    fi
}

# ---- 5. 创建目录结构 ----
setup_dirs() {
    info "创建目录结构..."
    mkdir -p images exports data
    info "目录创建完成"
}

# ---- 6. 生成配置文件 ----
setup_config() {
    if [[ -f "config.json" ]]; then
        warn "config.json 已存在，跳过生成"
        return
    fi
    info "生成默认配置文件 config.json..."
    cat > config.json << 'EOF'
{
    "host": "0.0.0.0",
    "port": 8765,
    "cors_origins": ["*"],
    "log_level": "INFO",
    "max_content_length": 20971520
}
EOF
    info "config.json 已生成"
}

# ---- 7. 创建 systemd 服务（可选）----
setup_systemd() {
    if [[ "$PKG" != "apt-get" && "$PKG" != "yum" ]]; then
        return
    fi
    echo ""
    read -p "是否安装 systemd 服务（开机自启）？[Y/n]: " -n 1 -r
    echo ""
    if [[ ! $REPLY =~ ^[Nn]$ ]]; then
        SERVICE_FILE="/etc/systemd/system/baoxiao.service"
        info "创建 systemd 服务文件: $SERVICE_FILE"
        sudo tee "$SERVICE_FILE" > /dev/null << 'SERVICE_EOF'
[Unit]
Description=报销费用填写工具
After=network.target

[Service]
Type=simple
User={{USER}}
WorkingDirectory={{CWD}}
ExecStart={{PY}} server_cloud.py
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
SERVICE_EOF

        # 替换占位符
        CURR_USER=$(whoami)
        CURR_PY=$PY
        CURR_CWD=$(pwd)
        sudo sed -i "s/{{USER}}/$CURR_USER/g; s|{{PY}}|$CURR_PY|g; s|{{CWD}}|$CURR_CWD|g" "$SERVICE_FILE"

        sudo systemctl daemon-reload
        sudo systemctl enable baoxiao
        info "systemd 服务已安装并设置开机自启"
        info "启动服务: sudo systemctl start baoxiao"
        info "停止服务: sudo systemctl stop baoxiao"
        info "查看日志: sudo journalctl -u baoxiao -f"
    else
        info "跳过 systemd 安装"
    fi
}

# ---- 8. 防火墙 ----
setup_firewall() {
    PORT=$(grep -o '"port": [0-9]*' config.json 2>/dev/null | awk '{print $2}' || echo 8765)
    if [[ "$PKG" == "apt-get" ]]; then
        if command -v ufw &> /dev/null; then
            echo ""
            read -p "是否开放防火墙端口 $PORT？[Y/n]: " -n 1 -r
            echo ""
            if [[ ! $REPLY =~ ^[Nn]$ ]]; then
                sudo ufw allow $PORT/tcp comment '报销费用填写工具'
                info "防火墙端口 $PORT 已开放"
            fi
        fi
    elif [[ "$PKG" == "yum" ]]; then
        if command -v firewall-cmd &> /dev/null; then
            echo ""
            read -p "是否开放防火墙端口 $PORT？[Y/n]: " -n 1 -r
            echo ""
            if [[ ! $REPLY =~ ^[Nn]$ ]]; then
                sudo firewall-cmd --permanent --add-port=$PORT/tcp
                sudo firewall-cmd --reload
                info "防火墙端口 $PORT 已开放"
            fi
        fi
    fi
}

# ---- 9. 启动测试 ----
start_test() {
    echo ""
    info "开始启动测试..."
    PORT=$(grep -o '"port": [0-9]*' config.json 2>/dev/null | awk '{print $2}' || echo 8765)
    $PY server_cloud.py &
    PID=$!
    sleep 3
    if kill -0 $PID 2>/dev/null; then
        info "✅ 服务启动成功！(PID: $PID)"
        kill $PID 2>/dev/null || true
        echo ""
        echo "============================================"
        echo "🎉 安装完成！"
        echo "============================================"
        echo ""
        echo "启动服务:"
        echo "  ./venv/bin/python server_cloud.py"
        echo ""
        echo "然后用浏览器打开:"
        echo "  http://服务器IP:$PORT"
        echo ""
    else
        error "服务启动失败，请检查日志"
    fi
}

# ================================================================
# 配置指南（安装完成后显示）
# =============================================================
show_guide() {
    PORT=$(grep -o '"port": [0-9]*' config.json 2>/dev/null | awk '{print $2}' || echo 8765)
    echo ""
    echo "============================================"
    echo "📋 配置指南"
    echo "============================================"
    echo ""
    echo "1️⃣  修改配置文件 config.json"
    echo "   - 端口默认 8765，可改为 80（需root）或 443"
    echo "   - cors_origins 默认 '*'（任意来源），生产环境改为你的域名"
    echo ""
    echo "2️⃣  开放服务器防火墙端口"
    echo "   阿里云控制台 → 安全组 → 入方向规则 → 添加:"
    echo "   协议: TCP | 端口: $PORT | 来源: 0.0.0.0/0"
    echo ""
    echo "3️⃣  （可选）绑定域名 + HTTPS"
    echo "   - 用 Nginx 反代到本服务"
    echo "   - 用 certbot 申请 Let's Encrypt 免费证书"
    echo ""
    echo "4️⃣  （可选）开机自启"
    echo "   systemctl enable baoxiao"
    echo ""
    echo "5️⃣  重启服务"
    echo "   ./venv/bin/python server_cloud.py"
    echo ""
    echo "📝 前端修改提示:"
    echo "   将 index.html 中的 API 地址从 localhost:8765"
    echo "   改为 http://你的服务器IP:$PORT"
    echo ""
}

# ================================================================
# 主流程
# ================================================================
main() {
    detect_os
    check_python
    install_system_deps
    install_python_deps
    setup_dirs
    setup_config

    echo ""
    echo "============================================"
    echo "可选配置（按 Enter 跳过）"
    echo "============================================"
    setup_systemd
    setup_firewall
    start_test
    show_guide
}

main "$@"
