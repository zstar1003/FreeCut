@echo off
echo ===============================================
echo FreeCut 插件安装脚本
echo ===============================================

REM 设置变量
set PLUGIN_NAME=FreeCut
set PLUGIN_PATH=%~dp0bin\Debug
set INSTALL_PATH=C:\FreeCut
set DLL_NAME=FreeCut.dll

echo.
echo 1. 检查构建文件...
if not exist "%PLUGIN_PATH%\%DLL_NAME%" (
    echo 错误: 找不到 %DLL_NAME%
    echo 请先构建项目: dotnet build FreeCut.csproj
    pause
    exit /b 1
)

echo 构建文件检查通过!

echo.
echo 2. 创建安装目录...
if not exist "%INSTALL_PATH%" mkdir "%INSTALL_PATH%"

echo.
echo 3. 复制插件文件...
copy "%PLUGIN_PATH%\*.dll" "%INSTALL_PATH%\" > nul
copy "%PLUGIN_PATH%\*.pdb" "%INSTALL_PATH%\" > nul 2>nul
copy "ribbon.xml" "%INSTALL_PATH%\" > nul 2>nul

echo 文件复制完成!

echo.
echo 4. 注册插件到PowerPoint...

REM 创建注册表文件
echo Windows Registry Editor Version 5.00 > temp_register.reg
echo. >> temp_register.reg
echo [HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\FreeCut.ThisAddIn] >> temp_register.reg
echo "Description"="FreeCut - PPT自动裁剪PDF导出插件" >> temp_register.reg
echo "FriendlyName"="FreeCut" >> temp_register.reg
echo "LoadBehavior"=dword:00000003 >> temp_register.reg
echo "Manifest"="file:///%INSTALL_PATH:\=/%/FreeCut.dll.manifest" >> temp_register.reg

REM 导入注册表
regedit /s temp_register.reg
del temp_register.reg

echo 插件注册完成!

echo.
echo 5. 创建清单文件...
echo ^<?xml version="1.0" encoding="utf-8"?^> > "%INSTALL_PATH%\FreeCut.dll.manifest"
echo ^<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0"^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"
echo   ^<assemblyIdentity name="FreeCut" version="1.0.0.0" type="win32" /^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"
echo   ^<file name="FreeCut.dll" hashalg="SHA1"^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"
echo     ^<comClass clsid="{12345678-1234-1234-1234-123456789ABC}" threadingModel="Both" /^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"
echo   ^</file^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"
echo ^</assembly^> >> "%INSTALL_PATH%\FreeCut.dll.manifest"

echo 清单文件创建完成!

echo.
echo ===============================================
echo 安装完成！
echo ===============================================
echo.
echo 安装位置: %INSTALL_PATH%
echo.
echo 下一步操作：
echo 1. 关闭所有PowerPoint窗口
echo 2. 重新启动PowerPoint
echo 3. 查看Ribbon中是否出现"FreeCut"标签页
echo.
echo 如果插件未显示，请检查：
echo - PowerPoint信任中心设置
echo - 宏安全级别设置
echo - 插件加载行为设置
echo.
echo 启动PowerPoint进行测试...
start "" "powerpnt.exe"

echo.
echo 按任意键退出...
pause > nul