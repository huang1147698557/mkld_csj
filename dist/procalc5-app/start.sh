#!/bin/bash
echo "========================================"
echo "  Procalc5 Rotor Calculator Automation"
echo "========================================"
echo

# 检查 Java 是否安装
if ! command -v java &> /dev/null; then
    echo "[错误] 未检测到 Java，请先安装 JDK 8 或以上版本。"
    echo "  brew install openjdk@17"
    echo "  或从 https://www.oracle.com/java/technologies/downloads/ 下载"
    read -p "按回车键退出..."
    exit 1
fi

# 脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# 默认工作目录（Mac/Linux 使用用户主目录下的 procalc5）
if [ -z "$WORKDIR" ]; then
    WORKDIR="$HOME/procalc5"
fi

# 确保工作目录存在
mkdir -p "$WORKDIR"

echo "工作目录: $WORKDIR"
echo "ChromeDriver: 由 Selenium Manager 自动管理（无需手动安装）"
echo

# 启动程序
java -jar "$SCRIPT_DIR/procalc5.jar" --workdir="$WORKDIR" "$@"

EXIT_CODE=$?
if [ $EXIT_CODE -ne 0 ]; then
    echo
    if echo "$OUTPUT" | grep -q "procalc5.proflute.xlsx"; then
        echo "[提示] 请将 procalc5.proflute.xlsx 放入 $WORKDIR 目录后重试"
    else
        echo "[提示] 程序异常退出 (错误码: $EXIT_CODE)"
    fi
    read -p "按回车键退出..."
fi
