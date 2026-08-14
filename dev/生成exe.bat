@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ==========================================
echo 草种测定管理 - EXE 启动器生成工具
echo ==========================================
echo.

set "CSC64=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
set "CSC32=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"

if exist "%CSC64%" (
    set "CSC=%CSC64%"
) else if exist "%CSC32%" (
    set "CSC=%CSC32%"
) else (
    echo [错误] 未找到 C# 编译器 csc.exe
    echo.
    pause
    exit /b 1
)

if not exist "launcher.cs" (
    echo [错误] 未找到 launcher.cs
    pause
    exit /b 1
)

if not exist "assets\app_icon.ico" (
    echo [错误] 未找到 assets\app_icon.ico
    pause
    exit /b 1
)

echo 正在生成 草种测定管理.exe ...
echo.

"%CSC%" ^
    /nologo ^
    /target:winexe ^
    /optimize+ ^
    /win32icon:"assets\app_icon.ico" ^
    /reference:System.Windows.Forms.dll ^
    /out:"草种测定管理.exe" ^
    "launcher.cs"

if errorlevel 1 (
    echo.
    echo [失败] EXE 生成失败。
    pause
    exit /b 1
)

echo.
echo ==========================================
echo 生成完成！
echo.
echo 草种测定管理.exe
echo ==========================================
echo.

pause