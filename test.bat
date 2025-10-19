@echo off
echo ===============================================
echo FreeCut 插件测试
echo ===============================================

echo.
echo 1. 编译测试程序...
csc TestFreeCut.cs /out:TestFreeCut.exe /target:exe

if %ERRORLEVEL% neq 0 (
    echo 编译失败！
    pause
    exit /b 1
)

echo 编译成功！

echo.
echo 2. 运行测试程序...
TestFreeCut.exe

echo.
echo 测试完成！

REM 清理临时文件
if exist TestFreeCut.exe del TestFreeCut.exe

pause