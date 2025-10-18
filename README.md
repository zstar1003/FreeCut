# FreeCut - PPT自动裁剪PDF导出插件

## 🎯 功能简介

FreeCut是一个PowerPoint插件，专为学术论文写作设计。它可以选中PPT页面后，自动裁剪并导出为PDF文件，解决了PPT导出PDF后需要手动裁剪的问题。

## ✨ 主要功能

- ✅ **智能页面选择**: 支持单页和多页选择导出
- ✅ **自动检测边界**: 智能识别内容区域，自动裁剪空白
- ✅ **自定义边距**: 支持上下左右边距精确设置（0-500像素）
- ✅ **高质量PDF导出**: 可配置DPI（72-600）和质量（1-100）
- ✅ **实时预览**: 所见即所得的裁剪效果预览
- ✅ **批量处理**: 一次性处理多个页面
- ✅ **设置持久化**: 自动保存用户偏好配置

## 🚀 当前版本：VSTO COM Add-in（推荐）

**✅ 专业的PowerPoint插件解决方案！**

我们基于VSTO（Visual Studio Tools for Office）技术开发了功能完整的PowerPoint COM Add-in，具有以下特性：

### 技术优势
- 🎯 **原生集成**: 直接集成到PowerPoint Ribbon界面
- 🔧 **高性能**: 基于.NET Framework 4.8，性能稳定
- 📦 **便捷安装**: 支持ClickOnce部署和注册表安装
- 🛡️ **安全可靠**: 通过数字签名，受信任执行
- 💻 **Windows原生**: 专为Windows PowerPoint优化

### 技术栈
- **后端**: C# .NET Framework 4.8
- **界面**: Windows Forms
- **PDF处理**: iTextSharp
- **Office集成**: Microsoft.Office.Interop.PowerPoint
- **部署**: VSTO部署或注册表安装

## 📥 安装和构建

### 1. 开发环境要求
- Windows 10/11
- Visual Studio 2017 或更新版本
- .NET Framework 4.8 Developer Pack
- PowerPoint 2016 或更新版本
- VSTO运行时（通常随Visual Studio安装）

### 2. 构建项目

```bash
# 1. 克隆项目
git clone [项目地址]
cd FreeCut

# 2. 使用Visual Studio构建
# 打开 FreeCut.csproj 并按 F6 构建

# 或使用MSBuild命令行
"C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe" FreeCut.csproj /p:Configuration=Release
```

### 3. 安装插件

#### 方法一：开发调试安装
1. 在Visual Studio中按F5运行项目
2. 会自动启动PowerPoint并加载插件

#### 方法二：注册表安装
1. 编译项目生成FreeCut.dll
2. 将DLL复制到合适位置
3. 运行以下注册表脚本：

```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\FreeCut.ThisAddIn]
"Description"="FreeCut - PPT自动裁剪PDF导出插件"
"FriendlyName"="FreeCut"
"LoadBehavior"=dword:00000003
"Manifest"="file:///C:/FreeCut/FreeCut.dll.manifest"
```

## 🎨 使用界面

### 主要功能区域

```
PowerPoint Ribbon → FreeCut标签页
├── 📄 PDF导出
│   ├── FreeCut设置 (打开完整设置面板)
│   └── 导出PDF (快速导出选中幻灯片)
├── 🔧 工具
│   ├── 预览裁剪 (查看裁剪效果)
│   └── 重新加载设置
└── ❓ 帮助
    └── 关于 (版本信息)
```

### 详细设置面板

```
FreeCut设置窗口
├── 📏 边距设置 (像素)
│   ├── 上边距、下边距、左边距、右边距 (0-500px)
│   └── 统一边距按钮
├── 🔍 自动检测设置
│   ├── ☑️ 自动检测边界
│   ├── 检测容差 (0-50)
│   └── 背景模式选择
├── 📤 导出设置
│   ├── PDF质量 (1-100)
│   ├── DPI选择 (72/150/300/600)
│   └── ☑️ 保持宽高比
└── 🚀 操作按钮
    ├── 保存设置、重置设置
    ├── 预览效果、导出PDF
    └── 取消
```

## ⚙️ 详细功能说明

### 裁剪设置
- **自动检测边界**:
  - ✅ 启用：智能识别内容区域，边距作为额外缓冲区
  - ❌ 禁用：按固定边距进行裁剪
- **边距控制**: 上下左右边距精确设置（0-500像素）
- **检测容差**: 自动检测时的颜色差异容忍度（0-50）
- **背景模式**: 自动检测/白色/透明/自定义颜色

### 导出设置
- **PDF质量**: 1-100，数值越高质量越好
- **导出DPI**:
  - 72 DPI：网页显示
  - 150 DPI：一般打印
  - 300 DPI：高质量印刷（推荐）
  - 600 DPI：专业印刷
- **保持宽高比**: 导出时保持原始页面比例

## 🎨 使用场景示例

### 📚 学术论文场景
```
任务：将PPT中的流程图导出到LaTeX论文中
操作步骤：
1. 在PowerPoint中选择包含流程图的幻灯片
2. 点击Ribbon中的"FreeCut设置"
3. 启用"自动检测边界"，设置边距为10像素
4. 选择300 DPI高质量导出
5. 点击"导出PDF"，选择保存位置
6. 在LaTeX中使用\includegraphics插入PDF
```

### 📊 研究报告场景
```
任务：批量导出图表制作报告附件
操作步骤：
1. 批量选择所有包含图表的幻灯片
2. 使用固定边距模式，设置统一的20px边距
3. 导出为一个PDF文件，每页图表尺寸一致
4. 直接作为研究报告的图表附件使用
```

## 🔧 项目结构

```
FreeCut/ (VSTO PowerPoint插件)
├── 📦 核心文件
│   ├── FreeCut.csproj           # 项目配置文件
│   ├── Properties/
│   │   └── AssemblyInfo.cs      # 程序集信息
│   └── ribbon.xml               # Ribbon界面定义
├── 💻 核心代码
│   ├── ThisAddIn.cs             # COM Add-in入口点
│   ├── FreeCutRibbon.cs         # Ribbon事件处理
│   ├── CropSettings.cs          # 设置管理
│   └── PdfExporter.cs           # PDF导出核心
├── 🖼️ 用户界面
│   ├── SettingsForm.cs          # 设置窗口
│   ├── ProgressForm.cs          # 进度显示
│   ├── PreviewForm.cs           # 预览窗口
│   └── *.Designer.cs            # 界面设计文件
└── 📚 文档和配置
    ├── README.md                # 本文档
    ├── .gitignore               # Git忽略配置
    └── *.resx                   # 资源文件
```

## 🐛 故障排除

### 常见问题

**❓ 插件未显示在Ribbon中**
- 检查PowerPoint版本（需要2016+）
- 确认.NET Framework 4.8已安装
- 检查VSTO运行时是否安装
- 验证注册表项是否正确

**❓ 构建失败**
- 确保安装了.NET Framework 4.8 Developer Pack
- 检查Visual Studio是否支持VSTO开发
- 验证NuGet包是否正确恢复

**❓ 导出功能异常**
- 检查PowerPoint是否选中了幻灯片
- 确认导出路径有写入权限
- 验证iTextSharp依赖是否正确加载

**❓ 权限和安全问题**
- 检查PowerPoint宏安全设置
- 确认插件是否被信任中心阻止
- 验证数字签名（如果使用）

## 📈 功能状态

### ✅ 已完成功能
- 完整的VSTO COM Add-in架构
- Ribbon界面集成
- 设置管理和持久化存储
- Windows Forms用户界面
- 基础PDF导出功能
- 图像裁剪算法框架
- 进度显示和预览功能

### 🚧 需要进一步开发
- 优化自动边界检测算法
- 完善PDF质量控制
- 增强图像处理性能
- 添加更多导出格式

### 💡 未来计划
- ClickOnce自动部署
- 多语言支持
- 云存储集成
- AI智能裁剪

## 📞 技术支持

### 构建环境配置
如遇到构建问题，请确保：
1. 安装Visual Studio 2017+（包含VSTO开发工具）
2. 安装.NET Framework 4.8 Developer Pack
3. 安装PowerPoint 2016+
4. 配置合适的代码签名证书（可选）

### 问题报告
如遇到问题，请提供：
- Windows版本和PowerPoint版本
- Visual Studio版本
- 详细的错误信息和操作步骤
- 示例PPT文件（如可能）

### 贡献代码
欢迎提交改进建议：
- 新功能需求
- 性能优化建议
- 用户体验改进
- 代码质量提升

## 📄 许可证

本软件遵循MIT许可证，允许自由使用、修改和分发。

---

## 💡 开发者说明

### 核心架构特点
1. **COM互操作**: 使用dynamic类型简化PowerPoint对象模型交互
2. **异步处理**: 避免UI阻塞，提供流畅的用户体验
3. **错误恢复**: 完善的异常处理和用户友好的错误提示
4. **模块化设计**: 清晰的职责分离，便于维护和扩展

### 性能优化
- 延迟加载UI组件
- 临时文件自动清理
- 内存使用优化
- 大文件处理优化

**🎉 FreeCut VSTO插件 - 专业的PowerPoint自动裁剪解决方案！🎉**