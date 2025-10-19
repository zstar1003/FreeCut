@echo off
echo ===============================================
echo FreeCut 插件卸载脚本
echo ===============================================

set INSTALL_PATH=C:\FreeCut

echo.
echo 1. 从PowerPoint中卸载插件...

REM 删除注册表项
reg delete "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\FreeCut.ThisAddIn" /f > nul 2>&1

echo 注册表清理完成!

echo.
echo 2. 删除插件文件...
if exist "%INSTALL_PATH%" (
    rmdir /s /q "%INSTALL_PATH%"
    echo 插件文件删除完成!
) else (
    echo 插件文件夹不存在，跳过删除。
)

echo.
echo ===============================================
echo 卸载完成！
echo ===============================================
echo.
echo 请重启PowerPoint以确保插件完全卸载。
echo.
pause