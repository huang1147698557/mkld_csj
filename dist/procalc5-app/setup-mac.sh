#!/bin/bash
echo "========================================"
echo "  Procalc5 - Mac Setup"
echo "========================================"
echo

WORKDIR="$HOME/procalc5"
mkdir -p "$WORKDIR"

echo "工作目录已创建: $WORKDIR"
echo

# 检查 chromedriver
if command -v chromedriver &> /dev/null; then
    CHROMEDRIVER_PATH=$(which chromedriver)
    echo "已检测到系统 chromedriver: $CHROMEDRIVER_PATH"
    echo "正在复制到工作目录..."
    cp "$CHROMEDRIVER_PATH" "$WORKDIR/chromedriver"
    echo "chromedriver 已复制到 $WORKDIR/chromedriver"
else
    echo "未检测到 chromedriver，推荐安装方式："
    echo "  方式一 (推荐): brew install chromedriver"
    echo "  方式二: 从 https://chromedriver.chromium.org/ 手动下载"
    echo
    read -p "是否现在通过 brew 安装 chromedriver? (y/n): " answer
    if [ "$answer" = "y" ] || [ "$answer" = "Y" ]; then
        brew install chromedriver
        if [ $? -eq 0 ]; then
            CHROMEDRIVER_PATH=$(which chromedriver)
            cp "$CHROMEDRIVER_PATH" "$WORKDIR/chromedriver"
            echo "chromedriver 已安装并复制到 $WORKDIR/chromedriver"
        fi
    fi
fi

echo
echo "请手动完成以下步骤:"
echo "  将 procalc5.proflute.xlsx 放入 $WORKDIR 目录"
echo
echo "运行程序:"
echo "  在终端执行: ./start.sh"
echo "  或双击 start.sh 文件"
