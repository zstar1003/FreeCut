# FreeCut 插件构建和安装指南

## 快速安装（推荐）

### 方法 1: 使用批处理安装脚本（最简单）

1. **在 Visual Studio 中构建项目**
   - 打开 `FreeCutPPT\PowerPointAddIn1\PowerPointAddIn1.sln`
   - 选择 `Release` 配置
   - 点击 `生成` → `重新生成解决方案`

2. **运行安装脚本**
   - 双击运行 `Install-FreeCut.bat`
   - 选择 `1` 安装插件
   - 等待安装完成

3. **重启 PowerPoint**
   - 在 Ribbon 中查找 `FreeCut` 标签页

### 方法 2: 使用 PowerShell 安装脚本

1. **在 Visual Studio 中构建项目**（同上）

2. **运行 PowerShell 脚本**
   - 右键点击 `Install-FreeCut.ps1`
   - 选择 `使用 PowerShell 运行`
   - 如果提示执行策略错误，运行：
     ```powershell
     Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
     ```
   - 选择 `1` 安装插件

3. **重启 PowerPoint**

---

## 详细说明

### 项目结构

```
FreeCut/
├── FreeCutPPT/
│   └── PowerPointAddIn1/              # VSTO 插件项目
│       ├── PowerPointAddIn1.csproj    # 项目文件
│       ├── bin/
│       │   ├── Debug/                 # 调试版本输出
│       │   └── Release/               # 发布版本输出
│       ├── PdfExporter.cs             # PDF 导出核心
│       ├── CropSettings.cs            # 设置管理
│       ├── SettingsForm.cs            # 设置界面
│       └── ...
├── Install-FreeCut.bat                # 批处理安装脚本 ⭐
├── Install-FreeCut.ps1                # PowerShell 安装脚本 ⭐
├── FreeCutInstaller.cs                # C# 安装器源码（需要编译）
└── README.md                          # 项目说明

⭐ = 推荐使用的安装脚本
```

### 在 Visual Studio 中构建

#### 步骤 1: 打开项目

```bash
# 使用 Visual Studio 打开项目文件
FreeCutPPT\PowerPointAddIn1\PowerPointAddIn1.sln
```

#### 步骤 2: 恢复 NuGet 包

Visual Studio 会自动恢复 NuGet 包，或者手动操作：
- 右键点击解决方案 → `还原 NuGet 程序包`

#### 步骤 3: 选择配置

- 顶部工具栏选择 `Release` 配置
- 平台选择 `Any CPU`

#### 步骤 4: 生成项目

- 菜单栏 → `生成` → `重新生成解决方案`
- 或按快捷键 `Ctrl+Shift+B`

#### 步骤 5: 检查输出

生成成功后，输出文件位于：
```
FreeCutPPT\PowerPointAddIn1\bin\Release\
├── PowerPointAddIn1.dll        # 主插件文件
├── PowerPointAddIn1.dll.manifest
├── SkiaSharp.dll               # PDF 生成库
├── System.Buffers.dll
├── System.Memory.dll
└── ...其他依赖文件
```

---

## 安装插件的三种方法

### 方法 1: 批处理脚本（推荐，最简单）

**优点：**
- ✅ 无需编译，双击即可运行
- ✅ 自动查找 DLL 文件
- ✅ 自动复制所有依赖
- ✅ 自动注册到 PowerPoint

**使用步骤：**
1. 双击 `Install-FreeCut.bat`
2. 输入 `1` 选择安装
3. 按提示操作

**卸载步骤：**
1. 双击 `Install-FreeCut.bat`
2. 输入 `2` 选择卸载

### 方法 2: PowerShell 脚本

**优点：**
- ✅ 功能与批处理相同
- ✅ 有彩色输出，更美观
- ✅ 错误处理更完善

**使用步骤：**
1. 右键点击 `Install-FreeCut.ps1` → `使用 PowerShell 运行`
2. 如果遇到执行策略错误：
   ```powershell
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   ```
3. 重新运行脚本

### 方法 3: 在 Visual Studio 中直接调试

**优点：**
- ✅ 适合开发调试
- ✅ 自动启动 PowerPoint

**使用步骤：**
1. 在 Visual Studio 中打开项目
2. 按 `F5` 启动调试
3. Visual Studio 会自动：
   - 构建项目
   - 注册插件
   - 启动 PowerPoint
   - 加载插件

---

## 常见问题

### Q1: 安装后在 PowerPoint 中看不到 FreeCut 标签页

**解决方案：**

1. **检查插件是否被禁用**
   - 打开 PowerPoint
   - `文件` → `选项` → `加载项`
   - 底部选择 `COM 加载项` → 点击 `转到`
   - 确保 `FreeCut` 已勾选

2. **检查信任中心设置**
   - `文件` → `选项` → `信任中心` → `信任中心设置`
   - `宏设置` → 选择"禁用所有宏，但显示通知"
   - `受信任位置` → 添加 `C:\FreeCut`

3. **清除 Office 缓存**
   - 关闭所有 PowerPoint 窗口
   - 删除缓存文件夹：`%localappdata%\Apps\2.0`
   - 重新运行安装脚本

4. **检查注册表**
   - 按 `Win+R` 输入 `regedit`
   - 导航到：`HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins`
   - 确认存在 `PowerPointAddIn1.ThisAddIn` 键
   - `LoadBehavior` 值应为 `3`

### Q2: 运行安装脚本时找不到 DLL 文件

**解决方案：**
- 确保已在 Visual Studio 中构建项目（`Release` 或 `Debug` 配置）
- 检查是否存在以下任一路径：
  - `FreeCutPPT\PowerPointAddIn1\bin\Release\PowerPointAddIn1.dll`
  - `FreeCutPPT\PowerPointAddIn1\bin\Debug\PowerPointAddIn1.dll`

### Q3: PowerShell 脚本无法运行（执行策略错误）

**错误信息：**
```
无法加载文件 Install-FreeCut.ps1，因为在此系统上禁止运行脚本。
```

**解决方案：**
以管理员身份打开 PowerShell，运行：
```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

或者直接使用批处理脚本（`Install-FreeCut.bat`），不受执行策略限制。

### Q4: 编译项目时提示缺少 SkiaSharp

**解决方案：**
1. 在 Visual Studio 中右键点击项目
2. 选择 `管理 NuGet 程序包`
3. 搜索并安装 `SkiaSharp`（版本 3.119.1 或更高）

### Q5: 如何完全卸载插件

**方法 1: 使用脚本卸载**
- 运行 `Install-FreeCut.bat` 或 `Install-FreeCut.ps1`
- 选择卸载选项

**方法 2: 手动卸载**
1. 删除注册表项：
   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn
   ```
2. 删除安装文件夹：`C:\FreeCut`
3. 重启 PowerPoint

---

## 开发调试

### 调试模式运行

1. 在 Visual Studio 中打开项目
2. 设置断点（如需要）
3. 按 `F5` 启动调试
4. PowerPoint 会自动启动并加载插件
5. 操作插件触发断点

### 查看日志

调试输出会显示在 Visual Studio 的 `输出` 窗口中。

### 修改代码后重新加载

1. 关闭 PowerPoint
2. 在 Visual Studio 中修改代码
3. 重新构建项目（`Ctrl+Shift+B`）
4. 按 `F5` 重新启动调试

---

## 技术支持

### 构建环境要求

- Windows 10/11
- Visual Studio 2017 或更高版本
- .NET Framework 4.8 Developer Pack
- PowerPoint 2016 或更高版本
- VSTO 运行时（通常随 Visual Studio 安装）

### 问题反馈

如遇到问题，请提供：
- Windows 版本和 PowerPoint 版本
- Visual Studio 版本
- 详细的错误信息和操作步骤
- 构建日志（如果构建失败）

---

## 许可证

本软件遵循 MIT 许可证，允许自由使用、修改和分发。
