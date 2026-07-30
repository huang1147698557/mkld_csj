@echo off
chcp 65001 >nul
echo ========================================
echo   Procalc5 Rotor Calculator Automation
echo ========================================
echo.

:: 检查 Java 是否安装
java -version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Java，请先安装 JDK 8 或以上版本。
    echo 下载地址: https://www.oracle.com/java/technologies/downloads/
    pause
    exit /b 1
)

:: 默认工作目录
if "%WORKDIR%"=="" set WORKDIR=C:\procalc5

:: 确保工作目录存在
if not exist "%WORKDIR%" mkdir "%WORKDIR%"

echo 工作目录: %WORKDIR%
echo ChromeDriver: 由 Selenium Manager 自动管理（无需手动安装）
echo.

:: 启动程序
java -jar "%~dp0procalc5.jar" --workdir="%WORKDIR%" %*

if %errorlevel% neq 0 (
    echo.
    echo [提示] 程序异常退出，请检查:
    echo   1. Chrome 浏览器是否已安装
    echo   2. procalc5.proflute.xlsx 是否存在于 %WORKDIR% 目录
    echo   3. 网络连接是否正常
    pause
)
