# FreeCut 插件故障排除指南

## 🔴 问题：插件显示"未加载" / "加载COM加载项时出现运行错误"

这是最常见的 VSTO 插件部署问题。请按以下步骤逐一排查：

---

## ✅ 第一步：检查 VSTO Runtime

###  问题症状
- 插件在加载项列表中显示"未加载"
- LoadBehavior 值变为 2（表示加载失败）

### 解决方案：安装 VSTO Runtime

**下载链接：**
- **官方下载页**：https://aka.ms/vs/17/release/vs_vsto.exe
- 或者搜索：`Microsoft Visual Studio 2010 Tools for Office Runtime`

**安装步骤：**
1. 下载 `vs_vsto.exe`
2. 以管理员身份运行
3. 完成安装后重启计算机
4. 重新安装 FreeCut 插件

**验证是否安装：**
- 打开"控制面板" → "程序和功能"
- 查找"Microsoft Visual Studio 2010 Tools for Office Runtime"

---

## ✅ 第二步：检查 .NET Framework 版本

### 要求
- .NET Framework 4.8 或更高版本

### 检查方法
```cmd
# 在命令提示符中运行
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release
```

### 下载安装
- 下载链接：https://dotnet.microsoft.com/download/dotnet-framework/net48

---

## ✅ 第三步：验证文件完整性

### 检查安装目录

打开 `C:\FreeCut\`，确保包含以下文件：

**必需文件清单：**
```
C:\FreeCut\
├── PowerPointAddIn1.dll           ✅ 主插件
├── PowerPointAddIn1.dll.manifest  ✅ 应用程序清单
├── PowerPointAddIn1.vsto          ✅ 部署清单（关键！）
├── SkiaSharp.dll                  ✅ PDF库
├── System.Buffers.dll
├── System.Memory.dll
├── System.Numerics.Vectors.dll
├── System.Runtime.CompilerServices.Unsafe.dll
├── libSkiaSharp.dll               ✅ 原生库
└── Microsoft.Office.Tools.Common.v4.0.Utilities.dll
```

**如果缺少文件：**
- 重新运行安装器
- 确保在 Visual Studio 中构建了 Release 版本

---

## ✅ 第四步：检查注册表配置

### 打开注册表编辑器
1. 按 `Win+R`，输入 `regedit`
2. 导航到：
   ```
   HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn
   ```

### 验证注册表值

应该包含以下键值：

| 名称 | 类型 | 数据 |
|------|------|------|
| **Description** | REG_SZ | FreeCut - PPT自动裁剪PDF导出插件 |
| **FriendlyName** | REG_SZ | FreeCut |
| **LoadBehavior** | REG_DWORD | 3 |
| **Manifest** | REG_SZ | `C:\FreeCut\PowerPointAddIn1.vsto\|vstolocal` |

**关键点：**
- ✅ `Manifest` 值必须指向 `.vsto` 文件
- ✅ 路径末尾必须有 `|vstolocal`
- ✅ `LoadBehavior` 应该是 3（如果是 2 表示加载失败）

### 修复注册表（如果错误）

创建一个 `.reg` 文件并双击运行：

```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn]
"Description"="FreeCut - PPT自动裁剪PDF导出插件"
"FriendlyName"="FreeCut"
"LoadBehavior"=dword:00000003
"Manifest"="C:\\FreeCut\\PowerPointAddIn1.vsto|vstolocal"
```

---

## ✅ 第五步：检查 Office 信任设置

### 打开信任中心
1. PowerPoint → `文件` → `选项` → `信任中心` → `信任中心设置`

### 检查以下设置：

#### 1. 受信任位置
- 点击左侧"受信任位置"
- 点击"添加新位置"
- 浏览到 `C:\FreeCut`
- 勾选"同时信任此位置的子文件夹"
- 点击"确定"

#### 2. 加载项设置
- 点击左侧"加载项"
- 确保**未**勾选"要求应用程序加载项由受信任的发布者签署"
- 或者：如果勾选了，需要为插件添加数字签名

---

## ✅ 第六步：清除 Office 缓存

### 原因
Office 会缓存加载项的状态，有时需要手动清除。

### 清除步骤
1. **关闭所有 Office 应用程序**

2. **删除缓存目录**：
   ```
   %localappdata%\Apps\2.0\
   ```
   - 按 `Win+R`
   - 输入 `%localappdata%\Apps\2.0\`
   - 删除整个 `2.0` 文件夹

3. **删除注册表中的 LoadBehavior**：
   ```cmd
   reg delete "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /v LoadBehavior /f
   ```

4. **重新运行安装器**

5. **重启 PowerPoint**

---

## ✅ 第七步：查看详细错误信息

### 启用 VSTO 日志

创建以下注册表键：

```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\VSTO\AdvancedLogging]
"EnablePipelineLogging"=dword:00000001
"EnableAddInLogging"=dword:00000001

[HKEY_CURRENT_USER\Software\Microsoft\VSTO\LogFilePath]
@="C:\\VSTOLogs"
```

### 查看日志

1. 创建目录 `C:\VSTOLogs`
2. 重启 PowerPoint
3. 尝试加载插件
4. 查看 `C:\VSTOLogs` 目录中的日志文件

---

## 🔧 快速修复脚本

将以下内容保存为 `FixFreeCut.bat` 并以管理员身份运行：

```batch
@echo off
echo 正在修复 FreeCut 插件...

REM 1. 删除缓存
echo 1. 清除 Office 缓存...
if exist "%localappdata%\Apps\2.0\" rmdir /s /q "%localappdata%\Apps\2.0\"

REM 2. 重置注册表
echo 2. 重置注册表配置...
reg delete "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /f >nul 2>&1

REM 3. 重新注册
echo 3. 重新注册插件...
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /v Description /t REG_SZ /d "FreeCut - PPT自动裁剪PDF导出插件" /f
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /v FriendlyName /t REG_SZ /d "FreeCut" /f
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /v LoadBehavior /t REG_DWORD /d 3 /f
reg add "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn" /v Manifest /t REG_SZ /d "C:\FreeCut\PowerPointAddIn1.vsto|vstolocal" /f

echo.
echo 修复完成！请重启 PowerPoint。
pause
```

---

## 📋 完整检查清单

按顺序检查以下项目：

- [ ] 已安装 VSTO Runtime (v10.0)
- [ ] 已安装 .NET Framework 4.8
- [ ] `C:\FreeCut\` 目录包含所有必需文件
- [ ] `PowerPointAddIn1.vsto` 文件存在
- [ ] 注册表 `Manifest` 值指向 `.vsto` 文件
- [ ] 注册表 `LoadBehavior` 值为 3
- [ ] `C:\FreeCut` 在信任位置列表中
- [ ] 已清除 Office 缓存
- [ ] 已重启 PowerPoint

---

## 🆘 仍然无法解决？

### 收集诊断信息

请提供以下信息：

1. **系统信息**：
   - Windows 版本
   - PowerPoint 版本
   - .NET Framework 版本

2. **文件列表**：
   ```cmd
   dir C:\FreeCut /b
   ```

3. **注册表值**：
   ```cmd
   reg query "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn"
   ```

4. **VSTO 日志**（如果已启用）：
   - `C:\VSTOLogs\` 目录中的日志文件内容

5. **错误截图**：
   - PowerPoint 加载项列表截图
   - 任何错误对话框截图

---

## 💡 常见错误和解决方案汇总

| 错误 | 原因 | 解决方案 |
|------|------|----------|
| **未加载** | 缺少 VSTO Runtime | 安装 VSTO Runtime |
| **LoadBehavior = 2** | 加载失败 | 检查文件完整性和依赖项 |
| **文件未找到** | 路径错误 | 验证注册表 Manifest 路径 |
| **访问被拒绝** | 权限问题 | 添加到信任位置 |
| **程序集加载失败** | 缺少依赖 DLL | 确保所有 DLL 都已复制 |

---

**按照本指南逐步排查，99% 的加载问题都可以解决！** 🎉
