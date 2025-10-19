@echo off
setlocal enabledelayedexpansion

echo ===============================================
echo FreeCut 完整构建和打包脚本
echo ===============================================

REM 设置变量
set PROJECT_NAME=FreeCut
set INSTALLER_NAME=FreeCutInstaller
set BUILD_CONFIG=Release

echo.
echo 1. 清理之前的构建文件...
if exist bin rmdir /s /q bin
if exist obj rmdir /s /q obj
if exist "!INSTALLER_NAME!.exe" del "!INSTALLER_NAME!.exe"

echo.
echo 2. 恢复 NuGet 包...
dotnet restore FreeCut.csproj
if !ERRORLEVEL! neq 0 (
    echo NuGet 包恢复失败！
    pause
    exit /b 1
)

echo.
echo 3. 构建 FreeCut 插件...
dotnet build FreeCut.csproj --configuration !BUILD_CONFIG! --no-restore
if !ERRORLEVEL! neq 0 (
    echo FreeCut 插件构建失败！
    pause
    exit /b 1
)

echo.
echo 构建成功！检查生成的文件：
dir bin\!BUILD_CONFIG! /b

echo.
echo 4. 编译安装器程序...
csc FreeCutInstaller.cs /target:winexe /out:FreeCutInstaller.exe /win32icon:icon.ico /reference:System.Windows.Forms.dll
if !ERRORLEVEL! neq 0 (
    echo 安装器编译失败！尝试不使用图标...
    csc FreeCutInstaller.cs /target:winexe /out:FreeCutInstaller.exe /reference:System.Windows.Forms.dll
    if !ERRORLEVEL! neq 0 (
        echo 安装器编译失败！
        pause
        exit /b 1
    )
)

echo.
echo ===============================================
echo 构建完成！
echo ===============================================
echo.
echo 生成的文件：
echo - FreeCut.dll          (插件主文件)
echo - FreeCutInstaller.exe (安装程序)
echo.
echo 使用说明：
echo 1. 运行 FreeCutInstaller.exe
echo 2. 点击"安装插件"按钮
echo 3. 重启 PowerPoint 查看 FreeCut 标签页
echo.

if exist FreeCutInstaller.exe (
    echo 是否现在运行安装器？
    set /p choice="输入 Y 运行安装器，或按任意键退出: "
    if /i "!choice!"=="Y" (
        echo 启动安装器...
        start FreeCutInstaller.exe
    )
) else (
    echo 错误：未找到 FreeCutInstaller.exe
    pause
)

echo.
echo 构建脚本执行完成！
pause