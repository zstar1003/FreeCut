# FreeCut 安装器 EXE 生成指南

## 🎯 目标
生成 `FreeCutInstaller.exe` 安装程序，方便用户一键安装 FreeCut 插件。

---

## 方法 1: 在 Visual Studio 中构建（推荐，最简单）

### 步骤 1: 打开安装器项目

1. 打开 Visual Studio
2. 选择 `文件` → `打开` → `项目/解决方案`
3. 浏览到 `E:\code\FreeCut\FreeCutInstaller\FreeCutInstaller.csproj`
4. 点击"打开"

### 步骤 2: 构建项目

1. 顶部工具栏选择 **Release** 配置
2. 点击 `生成` → `重新生成解决方案`
3. 或者按快捷键 `Ctrl+Shift+B`

### 步骤 3: 获取 EXE 文件

构建成功后，EXE 文件位于：
```
E:\code\FreeCut\FreeCutInstaller\bin\Release\FreeCutInstaller.exe
```

该 EXE 不是独立安装包，必须与下面列出的插件构建输出一起分发。

---

## 方法 2: 使用 Developer Command Prompt for Visual Studio

在仓库根目录运行：

```cmd
msbuild FreeCutInstaller\FreeCutInstaller.csproj /p:Configuration=Release /t:Rebuild
```

输出位于 `FreeCutInstaller\bin\Release\FreeCutInstaller.exe`。

---

## 📦 使用生成的 EXE 安装器

### 安装包内容

将以下文件打包在一起分发给用户：

```
FreeCut-Release/
├── FreeCutInstaller.exe                          # 安装器 EXE
└── FreeCutPPT/
    └── PowerPointAddIn1/
        └── bin/
            └── Release/
                ├── PowerPointAddIn1.dll          # 主插件
                ├── PowerPointAddIn1.vsto          # 部署清单（必需）
                ├── PowerPointAddIn1.dll.manifest
                ├── SkiaSharp.dll                 # 依赖库
                ├── System.Buffers.dll
                ├── System.Memory.dll
                ├── System.Numerics.Vectors.dll
                ├── System.Runtime.CompilerServices.Unsafe.dll
                └── libSkiaSharp.dll              # SkiaSharp 原生库
```

### 用户使用方法

1. 解压安装包
2. 双击运行 `FreeCutInstaller.exe`
3. 点击"安装插件"按钮
4. 等待安装完成
5. 重启 PowerPoint 即可看到 FreeCut 标签页

---

## 🔧 故障排除

### 问题 1: Visual Studio 中找不到项目类型

**原因：** 缺少 .NET 桌面开发工作负载

**解决：**
1. 打开 Visual Studio Installer
2. 点击"修改"
3. 勾选 `.NET 桌面开发`
4. 点击"修改"安装

### 问题 2: 构建脚本找不到 MSBuild

**原因：** 未安装 Visual Studio 或 MSBuild 不在标准路径

**解决：**
- 使用方法 1 在 Visual Studio 中构建
- 或者安装 Visual Studio Build Tools

### 问题 3: 生成的 EXE 无法运行

**原因：** 用户电脑缺少 .NET Framework 4.8

**解决：**
- 安装 .NET Framework 4.8 Runtime
- 下载地址：https://dotnet.microsoft.com/download/dotnet-framework/net48

---

## 📋 快速检查清单

### 构建前检查

- [ ] 已安装 Visual Studio 2017 或更高版本
- [ ] 已安装 .NET Framework 4.8 Developer Pack
- [ ] 已安装 .NET 桌面开发工作负载

### 构建步骤

1. [ ] 打开 Visual Studio
2. [ ] 打开 `FreeCutInstaller\FreeCutInstaller.csproj`
3. [ ] 选择 Release 配置
4. [ ] 点击"重新生成解决方案"
5. [ ] 检查 `bin\Release` 目录中的 EXE 文件

### 测试步骤

1. [ ] 确保已构建 PowerPointAddIn1 项目（Release 配置）
2. [ ] 运行 FreeCutInstaller.exe
3. [ ] 点击"安装插件"
4. [ ] 检查安装是否成功
5. [ ] 重启 PowerPoint 测试插件

---

## 💡 高级选项

### 添加应用程序图标

1. 将图标文件（icon.ico）放在 `FreeCutInstaller` 目录
2. 在项目属性中设置图标：
   - 右键项目 → 属性 → 应用程序
   - 图标 → 浏览 → 选择 icon.ico

### 签名 EXE 文件（可选）

为了避免 Windows SmartScreen 警告，可以对 EXE 进行数字签名：

```cmd
signtool sign /f certificate.pfx /p password /t http://timestamp.digicert.com FreeCutInstaller.exe
```

---

## 🚀 一键构建和打包脚本

创建完整的发布包：

```batch
@echo off
echo 正在构建 FreeCut 完整安装包...

REM 1. 构建插件
msbuild FreeCutPPT\PowerPointAddIn1\PowerPointAddIn1.csproj /p:Configuration=Release /t:Rebuild

REM 2. 构建安装器
msbuild FreeCutInstaller\FreeCutInstaller.csproj /p:Configuration=Release /t:Rebuild

REM 3. 创建发布目录
mkdir Release-Package
copy FreeCutInstaller\bin\Release\FreeCutInstaller.exe Release-Package\
xcopy /E /I FreeCutPPT\PowerPointAddIn1\bin\Release Release-Package\FreeCutPPT\PowerPointAddIn1\bin\Release\

echo.
echo 发布包已创建: Release-Package\
echo 可以将此目录打包为 ZIP 分发给用户
pause
```

---

## 📖 相关文档

- **加载故障排查**: [TROUBLESHOOTING.md](../TROUBLESHOOTING.md)
- **项目说明**: [README.md](../README.md)

---

**现在就在 Visual Studio 中打开项目，3 分钟生成 EXE！** 🎉

## 自动化回归检查

`.github/workflows/installer.yml` 在 Windows 上编译安装器，并运行 `tests/InstallPackageTests.cs`：覆盖缺失/空文件、不同工作目录、相同路径重装和覆盖复制。它不运行 PowerPoint，也不验证 Office 的加载/证书信任。

发布前仍需在 Windows + PowerPoint 手动验证：缺少 VSTO 时停止且不修改文件/注册表；分别删除两种清单时停止；完整包安装后实际出现标签页；关闭 PowerPoint 后重装与卸载；32/64 位 Office 支持范围。只有人工验证后才发布新的 ZIP。
