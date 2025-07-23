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

# 如果没有提供参数，列出当前目录中的PDF文件
if [ $# -lt 1 ]; then
    echo -e "${YELLOW}用法: $0 <pdf文件路径> [ppt输出路径]${NC}"
    echo -e "${YELLOW}列出当前目录中的PDF文件:${NC}"
    
    # 激活虚拟环境
    source "$VENV_DIR/bin/activate"
    
    # 检查简单版PDF转PPT工具是否存在
    if [ ! -f "$SCRIPT_DIR/simple_pdf2ppt.py" ]; then
        echo -e "${RED}错误: 工具文件不存在: simple_pdf2ppt.py${NC}"
        exit 1
    fi
    
    # 运行Python脚本列出PDF文件
    "$VENV_DIR/bin/python" "$SCRIPT_DIR/simple_pdf2ppt.py"
    
    exit 1
fi

# 获取输入文件路径
INPUT_PDF="$1"

# 获取输出文件路径（如果提供）
OUTPUT_PPT=""
if [ $# -ge 2 ]; then
    OUTPUT_PPT="$2"
fi

echo -e "${YELLOW}正在转换PDF到PPT...${NC}"

# 检查虚拟环境是否存在
if [ ! -d "$VENV_DIR" ]; then
    echo -e "${RED}错误: 虚拟环境不存在，请先设置环境${NC}"
    exit 1
fi

# 检查简单版PDF转PPT工具是否存在
if [ ! -f "$SCRIPT_DIR/simple_pdf2ppt.py" ]; then
    echo -e "${RED}错误: 工具文件不存在: simple_pdf2ppt.py${NC}"
    exit 1
fi

# 检查输入文件是否存在
if [ ! -f "$INPUT_PDF" ]; then
    echo -e "${RED}错误: 输入文件不存在: $INPUT_PDF${NC}"
    echo -e "${YELLOW}列出当前目录中的PDF文件:${NC}"
    
    # 激活虚拟环境
    source "$VENV_DIR/bin/activate"
    
    # 运行Python脚本列出PDF文件
    "$VENV_DIR/bin/python" "$SCRIPT_DIR/simple_pdf2ppt.py"
    
    exit 1
fi

# 激活虚拟环境并运行应用程序
source "$VENV_DIR/bin/activate"

# 检查必要的依赖
echo -e "${YELLOW}检查必要的依赖...${NC}"

# 检查pdf2pptx
if ! python -c "import pdf2pptx" &> /dev/null; then
    echo -e "${YELLOW}安装 pdf2pptx...${NC}"
    pip install pdf2pptx python-pptx
fi

# 运行转换
echo -e "${GREEN}开始转换...${NC}"

if [ -z "$OUTPUT_PPT" ]; then
    # 没有提供输出路径
    "$VENV_DIR/bin/python" "$SCRIPT_DIR/simple_pdf2ppt.py" "$INPUT_PDF"
else
    # 提供了输出路径
    "$VENV_DIR/bin/python" "$SCRIPT_DIR/simple_pdf2ppt.py" "$INPUT_PDF" "$OUTPUT_PPT"
fi

# 检查转换结果
if [ $? -eq 0 ]; then
    echo -e "${GREEN}转换完成!${NC}"
    exit 0
else
    echo -e "${RED}转换失败!${NC}"
    exit 1
fi 