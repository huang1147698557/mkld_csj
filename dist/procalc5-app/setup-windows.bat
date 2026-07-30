@echo off
chcp 65001 >nul
echo ========================================
echo   Procalc5 - Windows Setup
echo ========================================
echo.

set WORKDIR=C:\procalc5

:: 创建工作目录
if not exist "%WORKDIR%" mkdir "%WORKDIR%"

echo 工作目录已创建: %WORKDIR%
echo.
echo 请手动完成以下步骤:
echo.
echo 1. 下载 ChromeDriver:
echo    访问 https://chromedriver.chromium.org/downloads
echo    下载与您的 Chrome 版本匹配的 chromedriver.exe
echo    将 chromedriver.exe 放入 %WORKDIR% 目录
echo.
echo 2. 准备 Excel 输入文件:
echo    将 procalc5.proflute.xlsx 放入 %WORKDIR% 目录
echo.
echo 3. 运行程序:
echo    双击 start.bat 即可启动
echo.
pause
