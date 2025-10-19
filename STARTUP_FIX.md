# 🚨 FreeCut 启动问题解决方案

## 问题描述
如果您遇到以下错误：
```
无法启动程序"E:\code\FreeCut\bin\Release\FreeCut.dll"
FreeCut.dll 不是有效的 Win32 应用程序
```

## 🔍 问题原因
这是正常现象！FreeCut.dll是一个**插件库文件**，不是可执行程序，无法直接运行。

## ✅ 正确的测试和安装方法

### 方法1：使用自动安装脚本（推荐）

1. **构建项目**
   ```bash
   cd E:\code\FreeCut
   dotnet build FreeCut.csproj --configuration Debug
   ```

2. **运行安装脚本**
   ```bash
   # 右键点击 install.bat → "以管理员身份运行"
   install.bat
   ```

3. **测试插件**
   - 脚本会自动启动PowerPoint
   - 查看Ribbon中是否出现"FreeCut"标签页
   - 点击"FreeCut设置"测试功能

### 方法2：使用测试程序验证

1. **运行测试脚本**
   ```bash
   test.bat
   ```

2. **查看测试结果**
   ```
   ✓ DLL加载成功
   ✓ 找到 ThisAddIn 类
   ✓ 找到 CropSettings 类
   ✓ 找到 FreeCutRibbon 类
   ✓ 找到 PdfExporter 类
   ✓ ThisAddIn 类标记为COM可见
   ✓ ThisAddIn 类有GUID属性
   ```

### 方法3：手动安装（高级用户）

1. **复制文件**
   ```bash
   mkdir C:\FreeCut
   copy bin\Debug\*.dll C:\FreeCut\
   copy bin\Debug\*.pdb C:\FreeCut\
   ```

2. **注册插件**
   - 运行注册表编辑器 (regedit)
   - 导航到: `HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins`
   - 创建项: `FreeCut.ThisAddIn`
   - 添加以下值：
     ```
     Description = "FreeCut - PPT自动裁剪PDF导出插件"
     FriendlyName = "FreeCut"
     LoadBehavior = 3 (DWORD)
     Manifest = "file:///C:/FreeCut/FreeCut.dll.manifest"
     ```

## 🔧 调试配置（Visual Studio用户）

如果您使用Visual Studio进行开发，已经配置了调试启动：

1. **设置启动程序**：PowerPoint.exe
2. **F5调试**：会自动启动PowerPoint并加载插件
3. **断点调试**：可以在代码中设置断点进行调试

## ⚠️ 故障排除

### 插件未在PowerPoint中显示
1. **检查PowerPoint信任中心设置**
   - 文件 → 选项 → 信任中心 → 信任中心设置
   - 加载项 → 需要由受信任的发布者签署应用程序加载项（取消勾选）
   - 宏设置 → 启用所有宏

2. **检查插件加载状态**
   - PowerPoint → 文件 → 选项 → 加载项
   - 管理：COM加载项 → 转到
   - 查看FreeCut是否在列表中且已勾选

3. **重新注册插件**
   ```bash
   uninstall.bat  # 先卸载
   install.bat    # 重新安装
   ```

### 权限问题
- 以**管理员身份**运行安装脚本
- 确保C:\FreeCut目录有写入权限

### PowerPoint版本兼容性
- 支持PowerPoint 2016或更新版本
- 需要安装.NET Framework 4.7.2或更高版本

## 📁 文件说明

| 文件 | 用途 |
|------|------|
| `install.bat` | 自动安装插件到PowerPoint |
| `uninstall.bat` | 卸载插件 |
| `test.bat` | 测试DLL文件是否正确构建 |
| `TestFreeCut.cs` | 测试程序源代码 |
| `bin\Debug\FreeCut.dll` | 主插件文件 |
| `bin\Debug\FreeCut.pdb` | 调试符号文件 |

## 🎯 快速验证流程

```bash
# 1. 构建项目
dotnet build

# 2. 测试DLL
test.bat

# 3. 安装插件
install.bat

# 4. 启动PowerPoint验证
# 查看Ribbon中的FreeCut标签页
```

## 💡 成功标志

当您看到以下内容时，说明安装成功：
- ✅ PowerPoint Ribbon中出现"FreeCut"标签页
- ✅ 点击"FreeCut设置"能打开设置窗口
- ✅ 插件加载时显示"FreeCut插件已成功加载！"消息

---

**记住：DLL文件是库文件，不是可执行程序！请使用上述方法正确安装和测试插件。** 🎉