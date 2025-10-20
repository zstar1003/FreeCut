@echo off
setlocal enabledelayedexpansion

echo ===============================================
echo FreeCut 安装器构建脚本
echo ===============================================
echo.

REM 查找 MSBuild
set MSBUILD_PATH=

REM 尝试常见的 MSBuild 路径
for %%v in (2022 2019 2017) do (
    if exist "C:\Program Files\Microsoft Visual Studio\%%v\Community\MSBuild\Current\Bin\MSBuild.exe" (
        set "MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\%%v\Community\MSBuild\Current\Bin\MSBuild.exe"
        goto :found_msbuild
    )
    if exist "C:\Program Files\Microsoft Visual Studio\%%v\Professional\MSBuild\Current\Bin\MSBuild.exe" (
        set "MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\%%v\Professional\MSBuild\Current\Bin\MSBuild.exe"
        goto :found_msbuild
    )
    if exist "C:\Program Files\Microsoft Visual Studio\%%v\Enterprise\MSBuild\Current\Bin\MSBuild.exe" (
        set "MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\%%v\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
        goto :found_msbuild
    )
    if exist "C:\Program Files (x86)\Microsoft Visual Studio\%%v\Community\MSBuild\Current\Bin\MSBuild.exe" (
        set "MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\%%v\Community\MSBuild\Current\Bin\MSBuild.exe"
        goto :found_msbuild
    )
)

echo 错误: 未找到 MSBuild！
echo 请确保已安装 Visual Studio 2017 或更高版本。
echo.
echo 或者，在 Visual Studio 中打开 FreeCutInstaller.csproj 并手动构建。
pause
exit /b 1

:found_msbuild
echo 找到 MSBuild: !MSBUILD_PATH!
echo.

echo 1. 清理之前的构建...
if exist "FreeCutInstaller\bin" rmdir /s /q "FreeCutInstaller\bin"
if exist "FreeCutInstaller\obj" rmdir /s /q "FreeCutInstaller\obj"

echo.
echo 2. 构建安装器项目...
"!MSBUILD_PATH!" FreeCutInstaller\FreeCutInstaller.csproj /p:Configuration=Release /p:Platform=AnyCPU /t:Rebuild /v:m

if !ERRORLEVEL! neq 0 (
    echo.
    echo 构建失败！
    pause
    exit /b 1
)

echo.
echo ===============================================
echo 构建成功！
echo ===============================================
echo.
echo 生成的文件位于:
echo   FreeCutInstaller\bin\Release\FreeCutInstaller.exe
echo.

REM 复制到根目录方便使用
if exist "FreeCutInstaller\bin\Release\FreeCutInstaller.exe" (
    copy /Y "FreeCutInstaller\bin\Release\FreeCutInstaller.exe" "." >nul
    echo 已复制到根目录: FreeCutInstaller.exe
    echo.
)

echo 使用说明:
echo 1. 确保已构建 PowerPointAddIn1 项目
echo 2. 运行 FreeCutInstaller.exe
echo 3. 点击"安装插件"按钮
echo.

pause
