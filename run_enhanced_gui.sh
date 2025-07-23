#!/bin/bash

# 定义颜色
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # 无颜色

# 获取脚本所在目录
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# 虚拟环境目录
VENV_DIR="$SCRIPT_DIR/venv"

echo -e "${YELLOW}启动增强版PDF转PPT GUI界面...${NC}"

# 检查Python版本
PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
PYTHON_MAJOR=$(echo $PYTHON_VERSION | cut -d. -f1)
PYTHON_MINOR=$(echo $PYTHON_VERSION | cut -d. -f2)

if [ "$PYTHON_MAJOR" -lt 3 ] || ([ "$PYTHON_MAJOR" -eq 3 ] && [ "$PYTHON_MINOR" -lt 7 ]); then
    echo -e "${RED}错误: 需要Python 3.7或更高版本${NC}"
    echo -e "${RED}当前版本: $PYTHON_VERSION${NC}"
    exit 1
fi

# 检查核心转换模块是否存在
if [ ! -f "$SCRIPT_DIR/pdf_to_ppt.py" ]; then
    echo -e "${RED}错误: 核心转换模块不存在: pdf_to_ppt.py${NC}"
    exit 1
fi

# 检查增强版GUI是否存在
if [ ! -f "$SCRIPT_DIR/enhanced_gui.py" ]; then
    echo -e "${RED}错误: 增强版GUI文件不存在: enhanced_gui.py${NC}"
    exit 1
fi

# 直接安装必要的依赖（不使用虚拟环境）
echo -e "${YELLOW}安装必要的依赖...${NC}"

# 安装基本依赖
pip3 install pymupdf python-pptx pdf2pptx

# 尝试安装tkinterdnd2（用于拖放支持）
pip3 install tkinterdnd2

# 检查ImageMagick
if ! command -v convert &> /dev/null; then
    echo -e "${YELLOW}警告: ImageMagick未安装，某些转换方法将不可用${NC}"
    echo -e "${YELLOW}在macOS上，可以使用 'brew install imagemagick' 安装${NC}"
    echo -e "${YELLOW}在Ubuntu上，可以使用 'sudo apt install imagemagick' 安装${NC}"
fi

# 运行增强版GUI
echo -e "${GREEN}启动增强版GUI...${NC}"
python3 "$SCRIPT_DIR/enhanced_gui.py" 