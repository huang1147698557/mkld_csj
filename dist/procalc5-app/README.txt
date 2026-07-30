@echo off
chcp 65001 >nul
echo.
echo ========================================
echo   Procalc5 Rotor 计算自动化工具
echo ========================================
echo.
echo 【使用方法】
echo.
echo 一、首次使用:
echo    Windows: 双击 setup-windows.bat
echo    Mac:     在终端运行 ./setup-mac.sh
echo.
echo 二、日常运行:
echo    Windows: 双击 start.bat
echo    Mac:     在终端运行 ./start.sh
echo.
echo 【文件说明】
echo.
echo   procalc5.jar         - 主程序 (跨平台 JAR)
echo   start.bat            - Windows 启动脚本
echo   start.sh             - Mac/Linux 启动脚本
echo   setup-windows.bat    - Windows 环境初始化
echo   setup-mac.sh         - Mac 环境初始化
echo   README.txt           - 本说明文件
echo.
echo 【工作目录】
echo   Windows: C:\procalc5
echo   Mac/Linux: ~/procalc5 (用户主目录下的 procalc5)
echo.
echo 工作目录中需要以下文件:
echo   - chromedriver (Mac) 或 chromedriver.exe (Windows)
echo   - procalc5.proflute.xlsx (输入参数 Excel 文件)
echo.
echo 【高级参数】
echo.
echo   自定义工作目录:
echo     Windows: start.bat --workdir=D:\mywork
echo     Mac:     ./start.sh --workdir=/Users/you/mywork
echo.
echo   自定义 chromedriver 路径:
echo     --chromedriver=/path/to/chromedriver
echo.
echo 【账号信息】
echo   已内置账号: EXTCNJANZHA
echo   如需修改请编辑源码后重新打包
echo.
pause
