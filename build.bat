@echo off
echo ========================================
echo FreeCut Office Add-in Build Script
echo ========================================

echo.
echo 检查Node.js环境...
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 错误: 未找到Node.js。
    echo.
    echo 🛠️  安装Node.js步骤:
    echo 1. 访问 https://nodejs.org/
    echo 2. 下载并安装 LTS 版本
    echo 3. 重新运行此脚本
    echo.
    echo 📖 或者查看 OFFICE_ADDIN_GUIDE.md 了解其他方案
    pause
    exit /b 1
)

echo ✅ Node.js 已安装
node --version

echo.
echo 检查npm...
npm --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 错误: npm未找到
    pause
    exit /b 1
)

echo ✅ npm 已安装
npm --version

echo.
echo 安装依赖包...
npm install
if %errorlevel% neq 0 (
    echo ❌ 错误: 依赖包安装失败
    pause
    exit /b 1
)

echo.
echo 构建项目...
npm run build
if %errorlevel% neq 0 (
    echo ❌ 错误: 构建失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo ✅ 构建完成！
echo ========================================
echo.
echo 输出文件位于 'dist' 文件夹:
echo - taskpane.html (主界面)
echo - function-file/function-file.html (函数文件)
echo - manifest.xml (插件清单)
echo.
echo 🚀 下一步:
echo 1. 启动开发服务器: npm run dev-server
echo 2. 在PowerPoint中加载插件: npm run sideload
echo 3. 或手动安装manifest.xml到Office
echo.
echo 📖 详细说明请查看 OFFICE_ADDIN_GUIDE.md
echo.
pause