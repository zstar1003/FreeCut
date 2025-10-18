# FreeCut VSTO 安装和构建指南

## 🎯 概述

本指南详细介绍如何设置开发环境、构建项目以及安装FreeCut PowerPoint插件。

## 📋 系统要求

### 最低要求
- **操作系统**: Windows 10 (版本1903及以上) 或 Windows 11
- **Office**: Microsoft PowerPoint 2016 或更新版本
- **框架**: .NET Framework 4.8
- **内存**: 4GB RAM (推荐8GB)
- **存储**: 至少500MB可用空间

### 开发环境要求
- **Visual Studio**: 2017, 2019, 或 2022 (任意版本)
- **VSTO**: Visual Studio Tools for Office (通常随VS安装)
- **.NET Framework 4.8 Developer Pack**
- **PowerPoint 2016+** (用于测试)

## 🛠️ 开发环境设置

### 1. 安装Visual Studio

#### 选项A：完整版Visual Studio
```
推荐安装组件：
- .NET Framework 4.8 开发工具
- Office/SharePoint 开发工具
- NuGet 包管理器
- Git 工具 (可选)
```

#### 选项B：Visual Studio Build Tools (最小安装)
如果您只需要构建项目，可以安装VS Build Tools：

1. 下载 [Visual Studio Build Tools](https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022)
2. 安装时选择：
   - .NET Framework 4.8 targeting pack
   - MSBuild
   - NuGet targets and build tasks

### 2. 安装.NET Framework 4.8 Developer Pack

如果构建时出现目标框架错误，请安装：
- [.NET Framework 4.8 Developer Pack](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48)

### 3. 验证安装

打开命令提示符，运行：
```cmd
# 检查.NET Framework版本
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\" /v Release

# 查找MSBuild
where msbuild

# 如果上面找不到，尝试VS安装路径
dir "C:\Program Files*\Microsoft Visual Studio\*\*\MSBuild\Current\Bin\MSBuild.exe" /s
```

## 📥 获取和构建项目

### 1. 克隆项目

```bash
# 使用Git克隆
git clone https://github.com/your-username/FreeCut.git
cd FreeCut

# 或者下载ZIP并解压
```

### 2. 构建项目

#### 方法A：使用Visual Studio
1. 打开 `FreeCut.csproj`
2. 选择 **生成 → 重新生成解决方案** (或按 Ctrl+Shift+B)
3. 检查输出窗口确认构建成功

#### 方法B：使用MSBuild命令行
```cmd
# 进入项目目录
cd E:\code\FreeCut

# 使用MSBuild构建 (调整路径到您的MSBuild位置)
"C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe" FreeCut.csproj /p:Configuration=Release

# 或者使用VS2022路径
"C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" FreeCut.csproj /p:Configuration=Release

# 社区版路径示例
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" FreeCut.csproj /p:Configuration=Release
```

#### 方法C：使用dotnet命令 (如果支持)
```bash
# 注意：可能需要先安装.NET Framework targeting pack
dotnet build FreeCut.csproj --configuration Release
```

### 3. 验证构建输出

构建成功后，检查以下文件是否生成：
```
bin/
├── Debug/ (或 Release/)
    ├── FreeCut.dll          # 主程序集
    ├── FreeCut.dll.manifest # 清单文件
    ├── FreeCut.pdb         # 调试符号
    └── [依赖DLL文件]        # NuGet包依赖
```

## 🔧 安装插件

### 方法1：开发调试安装 (推荐开发者)

1. 在Visual Studio中设置为启动项目
2. 按 **F5** 启动调试
3. Visual Studio会自动：
   - 构建项目
   - 注册插件
   - 启动PowerPoint
   - 加载插件进行调试

### 方法2：注册表手动安装

#### 步骤1：准备文件
```cmd
# 创建安装目录
mkdir C:\FreeCut
copy bin\Release\*.* C:\FreeCut\
```

#### 步骤2：创建注册表文件
创建 `FreeCut-Install.reg` 文件：

```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\FreeCut.ThisAddIn]
"Description"="FreeCut - PPT自动裁剪PDF导出插件"
"FriendlyName"="FreeCut"
"LoadBehavior"=dword:00000003
"Manifest"="file:///C:/FreeCut/FreeCut.dll.manifest"
```

#### 步骤3：执行安装
1. 右键点击 `.reg` 文件
2. 选择 **合并**
3. 确认添加到注册表

#### 步骤4：重启PowerPoint
关闭并重新打开PowerPoint，检查Ribbon中是否出现FreeCut标签页。

### 方法3：PowerShell自动安装脚本

创建 `Install-FreeCut.ps1`：

```powershell
# FreeCut 自动安装脚本
param(
    [string]$InstallPath = "C:\FreeCut",
    [switch]$Uninstall
)

if ($Uninstall) {
    # 卸载
    Write-Host "卸载FreeCut插件..."
    Remove-Item "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\FreeCut.ThisAddIn" -ErrorAction SilentlyContinue
    Remove-Item $InstallPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "卸载完成"
    return
}

# 安装
Write-Host "安装FreeCut插件到: $InstallPath"

# 创建安装目录
New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null

# 复制文件
$sourceFiles = Get-ChildItem "bin\Release\*" -Include "*.dll", "*.manifest", "*.pdb"
foreach ($file in $sourceFiles) {
    Copy-Item $file.FullName $InstallPath -Force
    Write-Host "复制: $($file.Name)"
}

# 注册插件
$regPath = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\FreeCut.ThisAddIn"
New-Item -Path $regPath -Force | Out-Null
Set-ItemProperty -Path $regPath -Name "Description" -Value "FreeCut - PPT自动裁剪PDF导出插件"
Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "FreeCut"
Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $regPath -Name "Manifest" -Value "file:///$($InstallPath.Replace('\','/').Replace(':','%3A'))/FreeCut.dll.manifest"

Write-Host "安装完成！请重启PowerPoint查看FreeCut标签页。"
```

使用方法：
```powershell
# 安装
.\Install-FreeCut.ps1

# 安装到自定义路径
.\Install-FreeCut.ps1 -InstallPath "D:\MyPrograms\FreeCut"

# 卸载
.\Install-FreeCut.ps1 -Uninstall
```

## 🔍 故障排除

### 构建错误

#### 错误：找不到.NET Framework 4.8
```
解决方案：
1. 下载安装 .NET Framework 4.8 Developer Pack
2. 或修改项目文件改用 .NET Framework 4.7.2
```

#### 错误：MSBuild不是内部或外部命令
```
解决方案：
1. 安装Visual Studio或Build Tools
2. 使用完整路径调用MSBuild
3. 或添加MSBuild路径到系统PATH
```

#### 错误：NuGet包还原失败
```
解决方案：
1. 检查网络连接
2. 清理NuGet缓存：nuget locals all -clear
3. 删除packages文件夹重新还原
4. 尝试使用国内NuGet源
```

### 运行时错误

#### 插件未在Ribbon中显示
```
排查步骤：
1. 检查注册表项是否正确
2. 验证清单文件路径
3. 检查PowerPoint信任中心设置
4. 查看Windows事件日志
```

#### 加载时出现安全警告
```
解决方案：
1. 检查PowerPoint宏安全设置
2. 将插件目录添加到信任位置
3. 使用代码签名证书（生产环境）
```

#### 功能调用失败
```
排查步骤：
1. 检查PowerPoint版本兼容性
2. 验证依赖库是否加载
3. 查看DebugView输出
4. 检查.NET Framework版本
```

## 📋 开发工作流

### 推荐开发流程

1. **环境准备**
   ```cmd
   git clone [repository]
   cd FreeCut
   ```

2. **开发调试**
   ```cmd
   # 使用Visual Studio
   devenv FreeCut.csproj
   # 按F5启动调试
   ```

3. **测试验证**
   - 测试所有主要功能
   - 验证不同PowerPoint版本兼容性
   - 检查错误处理逻辑

4. **发布准备**
   ```cmd
   # 构建发布版本
   msbuild FreeCut.csproj /p:Configuration=Release

   # 创建安装包
   powershell -ExecutionPolicy Bypass -File Install-FreeCut.ps1
   ```

### 调试技巧

1. **使用Visual Studio调试器**
   - 设置断点调试代码逻辑
   - 使用即时窗口检查变量

2. **启用详细日志**
   - 使用Debug.WriteLine输出调试信息
   - 配置DebugView监控输出

3. **PowerPoint开发者工具**
   - 启用PowerPoint开发工具选项卡
   - 使用VBA编辑器测试功能

## 📦 部署选项

### 开发测试部署
- 注册表安装（当前文档描述的方法）
- 适合开发和内部测试

### 生产环境部署
- **ClickOnce部署**：适合企业内部分发
- **MSI安装包**：适合正式产品分发
- **应用商店**：通过Microsoft AppSource分发

### 企业部署
- **组策略部署**：域环境批量安装
- **SCCM部署**：企业软件管理系统
- **自动化脚本**：PowerShell/批处理脚本

## 🎯 下一步

完成安装后，您可以：

1. **开始使用**：在PowerPoint中体验FreeCut功能
2. **自定义开发**：修改代码适应特定需求
3. **贡献代码**：提交改进建议和bug修复
4. **创建部署包**：为其他用户创建安装程序

## 📞 获取帮助

如果遇到问题，请：

1. **查看日志**：检查Windows事件查看器中的错误
2. **搜索文档**：查看项目README和相关文档
3. **社区支持**：在项目Issues中提问
4. **联系开发者**：通过项目页面联系维护者

---

**🎉 祝您成功构建和安装FreeCut！如有问题随时寻求帮助。**