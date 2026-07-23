#!/bin/bash
# 片区分类开发工具 - 统信UOS/Ubuntu/Debian 启动脚本
set -e

echo "=== 片区分类开发工具 v1.3 ==="
echo "正在启动..."

PYTHON=""
if command -v python3 &> /dev/null; then
    PYTHON=python3
elif command -v python &> /dev/null; then
    PYTHON=python
else
    echo "错误: 未找到Python，请先安装Python 3"
    echo "  sudo apt install python3 python3-pip"
    exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "检查依赖..."
$PYTHON -c "import flask, openpyxl, requests" 2>/dev/null || {
    echo "正在安装依赖..."
    pip3 install --user flask openpyxl requests
}

PORT="${PORT:-5050}"
echo ""
echo "启动服务器: http://localhost:$PORT"
echo ""
echo "  功能说明:"
echo "  1. 打开浏览器访问 http://localhost:$PORT"
echo "  2. 上传Excel文件 (支持 .xlsx 格式)"
echo "  3. 选择公司名称列和地址列"
echo "  4. 点击地理编码 -> 片区分类 -> 导出结果"
echo "  5. 使用左侧工具栏绘制/编辑片区多边形"
echo ""
echo "按 Ctrl+C 停止服务器"
echo ""
$PYTHON app.py --port "$PORT"
