@echo off
setlocal
title Procalc5 Windows Installer Builder
chcp 65001 >nul

echo.
echo ========================================
echo   Procalc5 Windows Installer Builder
echo ========================================
echo.

set "JPACKAGE_PATH="
if not "%JAVA_HOME%"=="" if exist "%JAVA_HOME%\bin\jpackage.exe" set "JPACKAGE_PATH=%JAVA_HOME%\bin\jpackage.exe"
if "%JPACKAGE_PATH%"=="" for /r "%~dp0.tools" %%F in (jpackage.exe) do set "JPACKAGE_PATH=%%F"
if "%JPACKAGE_PATH%"=="" where jpackage.exe >nul 2>&1
if "%JPACKAGE_PATH%"=="" if errorlevel 1 (
    echo [ERROR] JDK 17 or newer with jpackage is required.
    echo         The current Java installation cannot create a Windows installer.
    echo         Install a current JDK, set JAVA_HOME, then open a new terminal.
    goto :failed
)

set "MAVEN_PATH="
for /d %%D in ("%~dp0.tools\maven\apache-maven-*") do if exist "%%~fD\bin\mvn.cmd" set "MAVEN_PATH=%%~fD\bin\mvn.cmd"
if "%MAVEN_PATH%"=="" where mvn.cmd >nul 2>&1
if "%MAVEN_PATH%"=="" if errorlevel 1 (
    echo [ERROR] Maven 3.8 or newer is required.
    echo         Install Maven and add its bin folder to PATH.
    goto :failed
)

if "%JAVAFX_JMODS%"=="" (
    if not exist "%~dp0javafx-sdk\javafx-jmods-17.0.2\javafx.controls.jmod" (
        if not exist "%~dp0javafx-sdk\javafx-jmods\javafx.controls.jmod" (
            echo [ERROR] JavaFX Windows jmods were not found.
            echo         Set JAVAFX_JMODS or place jmods under javafx-sdk.
            goto :failed
        )
    )
)

rem Pass -Type msi or -Type app-image as needed.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-windows.ps1" %*
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" goto :failed

echo.
echo Windows package build completed successfully.
pause
exit /b 0

:failed
echo.
echo Windows package build did not start or failed.
echo See WINDOWS-PACKAGING.md for the required tools.
pause
exit /b 1
