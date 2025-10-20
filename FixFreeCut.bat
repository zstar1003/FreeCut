@echo off
echo ===============================================
echo FreeCut 插件修复工具
echo ===============================================
echo.
echo 此工具将尝试修复插件加载问题。
echo 请确保已关闭所有 PowerPoint 窗口。
echo.
pause

echo.
echo 1. 清除 Office 缓存...
if exist "%localappdata%\Apps\2.0\" (
    rmdir /s /q "%localappdata%\Apps\2.0\"
    echo    缓存已清除
) else (
    echo    未找到缓存（可能已清除）
)
echo.

echo 2. 重置注册表配置...
reg delete "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /f >nul 2>&1
echo    注册表项已删除
echo.

echo 3. 重新注册插件...
set INSTALL_PATH=C:\FreeCut
set VSTO_PATH=%INSTALL_PATH%\PowerPointAddIn1.vsto

if not exist "%VSTO_PATH%" (
    echo    错误: 找不到 %VSTO_PATH%
    echo    请确保插件已正确安装到 C:\FreeCut
    pause
    exit /b 1
)

reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /v Description /t REG_SZ /d "FreeCut - PPT自动裁剪PDF导出插件" /f >nul
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /v FriendlyName /t REG_SZ /d "FreeCut" /f >nul
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /v LoadBehavior /t REG_DWORD /d 3 /f >nul
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /v Manifest /t REG_SZ /d "%VSTO_PATH%|vstolocal" /f >nul

echo    插件已重新注册
echo.

echo 4. 检查必需文件...
echo.
echo    检查以下文件是否存在：

set ALL_FILES_EXIST=1

if exist "%INSTALL_PATH%\PowerPointAddIn1.dll" (
    echo    [√] PowerPointAddIn1.dll
) else (
    echo    [X] PowerPointAddIn1.dll - 缺失！
    set ALL_FILES_EXIST=0
)

if exist "%INSTALL_PATH%\PowerPointAddIn1.dll.manifest" (
    echo    [√] PowerPointAddIn1.dll.manifest
) else (
    echo    [X] PowerPointAddIn1.dll.manifest - 缺失！
    set ALL_FILES_EXIST=0
)

if exist "%INSTALL_PATH%\PowerPointAddIn1.vsto" (
    echo    [√] PowerPointAddIn1.vsto
) else (
    echo    [X] PowerPointAddIn1.vsto - 缺失！
    set ALL_FILES_EXIST=0
)

if exist "%INSTALL_PATH%\SkiaSharp.dll" (
    echo    [√] SkiaSharp.dll
) else (
    echo    [X] SkiaSharp.dll - 缺失！
    set ALL_FILES_EXIST=0
)

echo.

if %ALL_FILES_EXIST%==0 (
    echo    警告: 某些必需文件缺失！
    echo    请重新运行 FreeCutInstaller.exe 进行完整安装。
    echo.
)

echo 5. 检查 VSTO Runtime...
reg query "HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" >nul 2>&1
if errorlevel 1 (
    reg query "HKLM\SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R" >nul 2>&1
    if errorlevel 1 (
        echo    [X] 未安装 VSTO Runtime！
        echo.
        echo    请下载并安装 VSTO Runtime:
        echo    https://aka.ms/vs/17/release/vs_vsto.exe
        echo.
    ) else (
        echo    [√] VSTO Runtime 已安装
    )
) else (
    echo    [√] VSTO Runtime 已安装
)

echo.
echo ===============================================
echo 修复完成！
echo ===============================================
echo.
echo 下一步操作：
echo 1. 如果缺少文件，请重新运行安装器
echo 2. 如果缺少 VSTO Runtime，请先安装它
echo 3. 重启 PowerPoint
echo 4. 在 PowerPoint 的"文件"→"选项"→"加载项"中检查插件状态
echo.
echo 如果问题仍然存在，请查看 TROUBLESHOOTING.md 文档。
echo.
pause
