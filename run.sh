#!/bin/bash

# 设置颜色输出
GREEN='\033[0;32m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# 检查 Python 是否安装
echo -e "${GREEN}检查 Python 环境...${NC}"
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}未找到 Python，请先安装 Python 3${NC}"
    exit 1
fi

# 检查 pip 是否安装
echo -e "${GREEN}检查 pip...${NC}"
if ! command -v pip &> /dev/null; then
    echo -e "${RED}未找到 pip，请先安装 pip${NC}"
    exit 1
fi

# 检查 requirements.txt 是否存在
echo -e "${GREEN}检查 requirements.txt...${NC}"
if [ ! -f "requirements.txt" ]; then
    echo -e "${RED}未找到 requirements.txt 文件${NC}"
    exit 1
fi

# 检查 main.py 是否存在
echo -e "${GREEN}检查 main.py...${NC}"
if [ ! -f "main.py" ]; then
    echo -e "${RED}未找到 main.py 文件${NC}"
    exit 1
fi

# 检查 data 目录是否存在，如果不存在则创建
echo -e "${GREEN}检查 data 目录...${NC}"
if [ ! -d "data" ]; then
    echo -e "${GREEN}创建 data 目录...${NC}"
    mkdir data
fi

# 检查 data 目录中是否有 Excel 文件
echo -e "${GREEN}检查 Excel 文件...${NC}"
if ! ls data/*.xls* 1> /dev/null 2>&1; then
    echo -e "${RED}data 目录下没有找到 Excel 文件${NC}"
    echo -e "${GREEN}请将 Excel 文件放入 data 目录中${NC}"
    exit 1
fi

# 安装依赖
echo -e "${GREEN}安装依赖包...${NC}"
pip install -r requirements.txt

# 如果安装成功，运行 Python 脚本
if [ $? -eq 0 ]; then
    echo -e "${GREEN}开始运行数据分析脚本...${NC}"
    python main.py
else
    echo -e "${RED}依赖包安装失败${NC}"
    exit 1
fi 