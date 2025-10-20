# FreeCut 项目结构

清理后的完整项目结构（已删除多余文件）

## 📁 根目录

```
FreeCut/
├── 📄 README.md                    # 项目主要说明文档
├── 📄 INSTALL.md                   # 详细安装指南
├── 📄 QUICKSTART.md                # 3分钟快速开始
├── 📄 TROUBLESHOOTING.md           # 故障排除完整指南
├── 🔧 build_installer.bat          # 构建安装器 EXE 的脚本
├── 🔧 FixFreeCut.bat               # 插件修复工具
├── 🔧 CleanupProject.bat           # 项目清理脚本（可选）
│
├── 📦 FreeCutInstaller/            # 安装器项目（Windows Forms）
│   ├── FreeCutInstaller.csproj    # 安装器项目文件
│   ├── Program.cs                  # 入口点
│   ├── InstallerForm.cs            # 主窗体
│   ├── Properties/                 # 项目属性
│   ├── BUILD.md                    # 安装器构建指南
│   └── UPDATE_NOTES.md             # 更新说明
│
└── 📦 FreeCutPPT/                  # 主插件项目
    └── PowerPointAddIn1/
        ├── PowerPointAddIn1.csproj # VSTO 插件项目
        ├── ThisAddIn.cs            # 插件入口点
        ├── FreeCutRibbon.cs        # Ribbon UI
        ├── PdfExporter.cs          # PDF 导出核心
        ├── CropSettings.cs         # 设置管理
        ├── SettingsForm.cs         # 设置窗体
        ├── PreviewForm.cs          # 预览窗体
        ├── ProgressForm.cs         # 进度窗体
        ├── Properties/             # 项目属性
        ├── packages/               # NuGet 包
        └── bin/                    # 构建输出
            └── Release/
                ├── PowerPointAddIn1.dll
                ├── PowerPointAddIn1.dll.manifest
                ├── PowerPointAddIn1.vsto
                └── ... (所有依赖 DLL)
```

## 📚 文档说明

### 主要文档

| 文件 | 用途 | 目标用户 |
|------|------|----------|
| **README.md** | 项目介绍、功能说明、技术栈 | 所有人 |
| **QUICKSTART.md** | 3分钟快速安装和使用 | 新用户 |
| **INSTALL.md** | 详细安装步骤和方法对比 | 用户/开发者 |
| **TROUBLESHOOTING.md** | 完整故障排除指南 | 遇到问题的用户 |

### 安装器文档

| 文件 | 用途 |
|------|------|
| **FreeCutInstaller/BUILD.md** | 如何在 VS 中构建安装器 EXE |
| **FreeCutInstaller/UPDATE_NOTES.md** | 安装器修复说明和技术细节 |

## 🔧 工具脚本

| 脚本 | 功能 | 使用场景 |
|------|------|----------|
| **build_installer.bat** | 自动构建安装器 EXE | 发布新版本 |
| **FixFreeCut.bat** | 修复插件加载问题 | 用户遇到"未加载"错误 |
| **CleanupProject.bat** | 清理项目多余文件 | 开发维护 |

## 🎯 关键文件

### 插件运行必需文件（在用户电脑上）

```
C:\FreeCut\
├── PowerPointAddIn1.dll           ✅ 主程序集
├── PowerPointAddIn1.dll.manifest  ✅ 应用程序清单
├── PowerPointAddIn1.vsto          ✅ 部署清单（关键！）
├── SkiaSharp.dll                  ✅ PDF 生成库
├── System.*.dll                   ✅ 依赖库
└── libSkiaSharp.dll               ✅ 原生库
```

### 构建输出文件

**插件 DLL** (Release 模式)：
```
FreeCutPPT/PowerPointAddIn1/bin/Release/
```

**安装器 EXE**：
```
FreeCutInstaller/bin/Release/FreeCutInstaller.exe
```

## 📦 发布清单

准备发布时需要打包的文件：

```
FreeCut-v1.0-Release/
├── FreeCutInstaller.exe               # 安装程序
├── README.md                          # 使用说明
├── QUICKSTART.md                      # 快速开始
├── TROUBLESHOOTING.md                 # 故障排除
├── FixFreeCut.bat                     # 修复工具
└── FreeCutPPT/
    └── PowerPointAddIn1/
        └── bin/
            └── Release/               # 所有插件文件
```

## 🗑️ 已删除的文件

以下文件已在清理过程中删除（多余或过时）：

### 旧的安装脚本
- ❌ FreeCutInstaller.cs（单文件安装器，已被项目替代）
- ❌ Install-FreeCut.bat
- ❌ Install-FreeCut.ps1
- ❌ install.bat
- ❌ install.ps1
- ❌ install_simple.bat
- ❌ uninstall.bat
- ❌ 安装文件说明.md

### 旧的构建脚本
- ❌ build.bat
- ❌ build_all.bat

### 测试和临时文件
- ❌ test.bat
- ❌ unblock_and_enable.bat
- ❌ convert_to_ico.ps1
- ❌ ConvertSvgToIco.cs
- ❌ TestFreeCut.cs
- ❌ TestFreeCutDll.cs

### 重复的源代码（根目录）
- ❌ App.xaml.cs
- ❌ CropSettings.cs
- ❌ FreeCutRibbon.cs
- ❌ IDTExtensibility2.cs
- ❌ MainWindow.xaml.cs
- ❌ PdfExporter.cs
- ❌ PreviewForm.cs / PreviewForm.Designer.cs
- ❌ ProgressForm.cs / ProgressForm.Designer.cs
- ❌ SettingsForm.cs / SettingsForm.Designer.cs
- ❌ ThisAddIn.cs

### 过时的文档
- ❌ BUILD_INSTALL_GUIDE.md
- ❌ DEVELOPMENT_GUIDE.md
- ❌ FINAL_SUMMARY.md
- ❌ OFFICE_ADDIN_GUIDE.md
- ❌ PROJECT_SUMMARY.md
- ❌ STARTUP_FIX.md
- ❌ SUCCESS_SUMMARY.md
- ❌ TEST_GUIDE.md

## ✅ 清理后的优势

1. **结构清晰** - 只保留必要的文件
2. **易于维护** - 源代码都在项目目录中
3. **用户友好** - 清晰的文档和工具
4. **易于发布** - 明确的发布清单

## 🚀 快速命令

### 开发
```bash
# 构建插件
msbuild FreeCutPPT\PowerPointAddIn1\PowerPointAddIn1.csproj /p:Configuration=Release

# 构建安装器
build_installer.bat
```

### 调试
```bash
# 在 Visual Studio 中
F5 启动调试
```

### 发布
```bash
# 1. 构建两个项目（Release 模式）
# 2. 复制 FreeCutInstaller.exe
# 3. 复制整个 bin\Release 目录
# 4. 打包成 ZIP
```

---

**项目现在结构清晰，准备好发布了！** 🎉
