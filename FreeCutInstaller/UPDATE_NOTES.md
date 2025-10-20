# FreeCut 安装器更新说明

## 🔧 已修复的问题

之前的安装器存在以下问题，导致插件显示"未加载"：

### 问题 1: 错误的清单文件
- ❌ **旧方式**: 手动创建简化的 `.dll.manifest` 文件
- ✅ **新方式**: 使用 Visual Studio 自动生成的完整清单文件

### 问题 2: 错误的注册表配置
- ❌ **旧方式**: 注册表指向 `PowerPointAddIn1.dll.manifest`
- ✅ **新方式**: 注册表指向 `PowerPointAddIn1.vsto`（VSTO 部署清单）

### 问题 3: 缺少 VSTO Runtime 检查
- ❌ **旧方式**: 没有检查必需的运行时组件
- ✅ **新方式**: 安装前检查并提示用户安装 VSTO Runtime

---

## 📦 新的安装器功能

### 1. 自动检查依赖项
- 检测 VSTO Runtime 是否已安装
- 如果缺失，提示用户下载链接

### 2. 正确的文件复制
安装器现在会复制所有必需文件：
```
C:\FreeCut\
├── PowerPointAddIn1.dll           # 主插件
├── PowerPointAddIn1.dll.manifest  # 应用程序清单（VS生成）
├── PowerPointAddIn1.vsto          # 部署清单（关键！）
├── SkiaSharp.dll                  # 依赖库
├── System.*.dll                   # 其他依赖
└── libSkiaSharp.dll               # 原生库
```

### 3. 正确的注册表配置
```
Manifest = C:\FreeCut\PowerPointAddIn1.vsto|vstolocal
```
注意末尾的 `|vstolocal` 标志。

---

## 🔄 如何使用新安装器

### 步骤 1: 重新构建安装器

在 Visual Studio 中：
1. 打开 `FreeCutInstaller\FreeCutInstaller.csproj`
2. 选择 `Release` 配置
3. 按 `Ctrl+Shift+B` 构建
4. 生成的 EXE 位于：`FreeCutInstaller\bin\Release\FreeCutInstaller.exe`

或运行构建脚本：
```cmd
build_installer.bat
```

### 步骤 2: 分发给用户

打包以下内容：
```
FreeCut-安装包/
├── FreeCutInstaller.exe               # 新的安装器
├── FreeCutPPT/
│   └── PowerPointAddIn1/
│       └── bin/
│           └── Release/
│               ├── PowerPointAddIn1.dll
│               ├── PowerPointAddIn1.dll.manifest
│               ├── PowerPointAddIn1.vsto
│               └── ... (所有 DLL)
├── FixFreeCut.bat                     # 修复工具（可选）
├── TROUBLESHOOTING.md                 # 故障排除指南
└── README.md                          # 使用说明
```

### 步骤 3: 用户安装流程

1. **运行安装器**: 双击 `FreeCutInstaller.exe`
2. **检查依赖**: 安装器会自动检查 VSTO Runtime
3. **如果缺失**: 提示用户下载并安装 VSTO Runtime
   - 下载地址：https://aka.ms/vs/17/release/vs_vsto.exe
4. **完成安装**: 点击"安装插件"
5. **重启 PowerPoint**: 查看 FreeCut 标签页

---

## 🛠️ 如果仍然出现"未加载"问题

### 快速修复

运行 `FixFreeCut.bat` 修复脚本，它会：
1. 清除 Office 缓存
2. 重置注册表配置
3. 重新注册插件
4. 检查必需文件
5. 验证 VSTO Runtime

### 详细排查

查看 `TROUBLESHOOTING.md` 文档，包含：
- 7 个详细的故障排除步骤
- 完整的检查清单
- 常见错误和解决方案
- 如何启用 VSTO 日志

---

## 📝 技术细节

### VSTO 部署模型

VSTO 插件使用两层清单结构：

1. **部署清单 (.vsto)**
   - 包含部署信息
   - 指向应用程序清单
   - 注册表应该指向这个文件

2. **应用程序清单 (.dll.manifest)**
   - 包含程序集依赖
   - 包含数字签名
   - 包含 VSTO 配置

### 正确的注册表格式

```
Manifest = [绝对路径]\PowerPointAddIn1.vsto|vstolocal
```

**关键点：**
- 必须是绝对路径（不能是 file:/// URL）
- 必须指向 `.vsto` 文件（不是 `.dll.manifest`）
- 必须以 `|vstolocal` 结尾

---

## ✅ 验证安装成功

### 检查注册表
```cmd
reg query "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn"
```

应该看到：
```
Manifest    REG_SZ    C:\FreeCut\PowerPointAddIn1.vsto|vstolocal
LoadBehavior    REG_DWORD    0x3
```

### 检查文件
```cmd
dir C:\FreeCut
```

应该看到 `PowerPointAddIn1.vsto` 文件。

### 在 PowerPoint 中检查
1. `文件` → `选项` → `加载项`
2. 底部选择"COM 加载项" → `转到`
3. 应该看到 `FreeCut` 并且已勾选
4. 状态应该显示"已加载"

---

## 🎉 现在安装器应该可以正常工作了！

关键改进：
- ✅ 使用正确的 VSTO 清单文件
- ✅ 正确的注册表配置
- ✅ 自动检查依赖项
- ✅ 提供修复工具和详细文档

如果用户仍然遇到问题，请让他们：
1. 运行 `FixFreeCut.bat`
2. 查看 `TROUBLESHOOTING.md`
3. 提供诊断信息
