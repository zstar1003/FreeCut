# FreeCut 安装器更新说明

## 安装前置检查修复

- 缺少 VSTO Runtime 时停止安装，不再允许忽略必需依赖后显示成功。
- 在复制和注册前检查 DLL、应用清单及部署清单存在且非空；复制后再次检查。
- 相对于安装器自身目录查找插件，避免从快捷方式启动时找不到文件。
- 安装器状态更新不再处理输入事件，避免安装/卸载过程被重复点击打断。
- “文件复制和注册完成”不代表已在 PowerPoint 中验证加载。Issue #1 尚无环境/错误详情，这些是确认的安装器缺陷修复，而非对该用户根因的确认。

## 🔧 早期修复说明

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
msbuild FreeCutInstaller\FreeCutInstaller.csproj /p:Configuration=Release /t:Rebuild
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

参阅 [TROUBLESHOOTING.md](../TROUBLESHOOTING.md)，先收集安装器状态、Office 版本与加载错误。仓库没有自动修复脚本，不要盲目清除 Office 缓存或绕过安全策略。

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
- 本安装器使用绝对文件路径
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
- ✅ 提供加载排查文档

如果用户仍然遇到问题，请让他们：
1. 查看 [TROUBLESHOOTING.md](../TROUBLESHOOTING.md)
2. 提供诊断信息
