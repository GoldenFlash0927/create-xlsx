#!/bin/bash
# Excel文件更新脚本

# 进入脚本所在目录
cd "$(dirname "$0")"

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "❌ 错误: 未找到 python3，请先安装 Python 3"
    exit 1
fi

# 检查并安装依赖
if [ ! -d "venv" ]; then
    echo "正在创建虚拟环境..."
    python3 -m venv venv
fi

echo "正在激活虚拟环境并安装依赖..."
source venv/bin/activate
pip install -q -r requirements.txt

# 运行Python脚本
echo "正在执行Excel更新脚本..."
python3 update_excel.py

# 退出虚拟环境
deactivate

