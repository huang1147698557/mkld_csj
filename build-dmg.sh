#!/bin/bash
# Procalc5 转子计算器 - Mac DMG 打包脚本
set -e

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$PROJECT_DIR"

APP_NAME="Procalc5"
APP_VERSION="1.0.0"
FAT_JAR="procalc5-app.jar"
MAIN_CLASS="com.sd.discovery.single.Procalc5App"
OUTPUT_DIR="$PROJECT_DIR/dist"

echo "===== Step 1: Maven 构建 fat JAR ====="
mvn clean package -Pjavafx-app -DskipTests -q
echo "Maven 构建完成"

# 验证 fat JAR
FAT_JAR_PATH="$PROJECT_DIR/target/$FAT_JAR"
if [ ! -f "$FAT_JAR_PATH" ]; then
    echo "ERROR: fat JAR 不存在: $FAT_JAR_PATH"
    exit 1
fi
echo "Fat JAR 大小: $(ls -lh $FAT_JAR_PATH | awk '{print $5}')"

echo ""
echo "===== Step 2: 准备 jpackage 输入目录 ====="
INPUT_DIR="$PROJECT_DIR/target/jpackage-input"
rm -rf "$INPUT_DIR"
mkdir -p "$INPUT_DIR"
cp "$FAT_JAR_PATH" "$INPUT_DIR/"
echo "输入目录: $INPUT_DIR"

echo ""
echo "===== Step 3: 生成 DMG 安装包 ====="
rm -rf "$OUTPUT_DIR/$APP_NAME.app" "$OUTPUT_DIR/$APP_NAME-$APP_VERSION.dmg"
mkdir -p "$OUTPUT_DIR"

jpackage \
  --name "$APP_NAME" \
  --input "$INPUT_DIR" \
  --main-jar "$FAT_JAR" \
  --main-class "$MAIN_CLASS" \
  --type dmg \
  --app-version "$APP_VERSION" \
  --dest "$OUTPUT_DIR" \
  --mac-package-name "$APP_NAME" \
  --mac-package-identifier "com.sd.procalc5" \
  --java-options "--add-opens=javafx.graphics/com.sun.javafx.tk=ALL-UNNAMED" \
  --java-options "--add-opens=javafx.controls/javafx.scene.control=ALL-UNNAMED" \
  --java-options "-Dfile.encoding=UTF-8"

echo ""
echo "===== 打包完成 ====="
ls -lh "$OUTPUT_DIR"/*.dmg 2>/dev/null || echo "未找到 DMG 文件"
echo "输出目录: $OUTPUT_DIR"
