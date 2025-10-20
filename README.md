# FreeCut - PPT自动裁剪PDF导出插件

## 🎯 功能简介

FreeCut是一个PowerPoint插件，专为学术论文写作设计。它可以选中PPT页面后，自动裁剪并导出为PDF文件，解决了PPT导出PDF后需要手动裁剪的问题。

## ✨ 主要功能

- ✅ **智能页面选择**: 支持单页和多页选择导出
- ✅ **自动检测边界**: 智能识别内容区域，自动裁剪空白
- ✅ **自定义边距**: 支持上下左右边距精确设置（0-500像素）
- ✅ **高质量PNG导出**: 可配置DPI（72-600），PDF功能开发中
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
- **后端**: C# .NET Framework 4.7.2
- **界面**: Windows Forms
- **图像处理**: System.Drawing (PNG导出)
- **PDF处理**: 开发中 (依赖库兼容性问题)
- **Office集成**: Microsoft.Office.Interop.PowerPoint
- **部署**: 可执行安装器或注册表安装

## ⚠️ PDF功能开发状态

### 当前版本
- ✅ **PNG导出功能**: 完全可用，支持高质量图像导出
- 🔄 **PDF导出功能**: 正在开发中，遇到依赖库兼容性问题

### PDF功能技术问题
由于 .NET Framework 4.7.2 环境下的依赖库兼容性问题：
- iTextSharp 5.5.13.3 包依赖解析失败
- PDFsharp 在当前构建环境下无法正确加载
- .NET Core/5.0+ 工具链与传统 .NET Framework 项目存在集成困难

### 临时解决方案
用户可以使用以下方式获得PDF输出：
1. **使用PNG导出 + 转换工具**:
   - 导出高质量PNG图片
   - 使用Adobe Acrobat、在线工具或其他软件转换为PDF
2. **PowerPoint原生导出**:
   - 使用PowerPoint的"导出为PDF"功能
   - 然后使用FreeCut对PDF进行二次裁剪处理

### PDF功能恢复计划
1. **环境升级**: 考虑升级到.NET Framework 4.8或.NET 6+
2. **依赖库替换**: 评估其他.NET Framework兼容的PDF库
3. **Visual Studio NuGet**: 使用Visual Studio包管理器手动安装
4. **分离架构**: 将PDF生成独立为单独的服务组件

## 📥 安装和构建

### 1. 开发环境要求
- Windows 10/11
- Visual Studio 2017 或更新版本
- .NET Framework 4.7.2 Developer Pack
- PowerPoint 2016 或更新版本
- VSTO运行时（通常随Visual Studio安装）

### 2. 构建项目

```bash
# 1. 克隆项目
git clone [项目地址]
cd FreeCut

# 2. 构建FreeCut插件
dotnet build FreeCut.csproj --configuration Release

# 3. 构建安装器
csc FreeCutInstaller.cs /target:winexe /out:FreeCutInstaller.exe /reference:System.Windows.Forms.dll

# 4. 或使用一键构建脚本
.\build_all.bat
```

### 3. 安装插件

#### 方法一：使用可执行安装器（推荐）
1. 构建或下载 `FreeCutInstaller.exe`
2. 以管理员权限运行安装器
3. 点击"安装插件"按钮，安装器会自动：
   - 检查并复制插件文件
   - 创建清单文件
   - 注册插件到PowerPoint
4. 重启PowerPoint，在Ribbon中查找"FreeCut"标签页

#### 方法二：开发调试安装
1. 在Visual Studio中按F5运行项目
2. 会自动启动PowerPoint并加载插件

#### 方法三：手动注册表安装
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
│   ├── DPI选择 (72/150/300/600)
│   └── ☑️ 保持宽高比
└── 🚀 操作按钮
    ├── 保存设置、重置设置
    ├── 预览效果、导出图片
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
- **导出DPI**:
  - 72 DPI：网页显示
  - 150 DPI：一般打印（默认）
  - 300 DPI：高质量印刷
  - 600 DPI：专业印刷
- **保持宽高比**: 导出时保持原始页面比例
- **PDF质量**: 控制 PDF 中图片的压缩质量（1-100）

## 🔬 技术原理：DPI 与 PDF 质量

### 什么是 DPI？

**DPI（Dots Per Inch，每英寸点数）** 是衡量图像分辨率的单位，表示每英寸包含多少个像素点。

#### DPI 与图像质量的关系

假设一张 PPT 幻灯片的物理尺寸为 **10 英寸 × 7.5 英寸**（标准 16:9 比例）：

| DPI | 输出像素尺寸 | 文件大小 | 适用场景 |
|-----|------------|---------|---------|
| **72 DPI** | 720 × 540 px | ~100 KB | 网页显示、屏幕浏览 |
| **150 DPI** | 1500 × 1125 px | ~500 KB | 一般打印、日常文档（默认） |
| **300 DPI** | 3000 × 2250 px | ~2 MB | 高质量印刷、学术论文 |
| **600 DPI** | 6000 × 4500 px | ~8 MB | 专业印刷、出版物 |

#### 为什么不同 DPI 下 PDF 页面大小相同？

**关键点**：FreeCut 使用 **DPI 自适应边距缩放** 技术，确保不同 DPI 下导出的 PDF 具有相同的物理尺寸。

```
原理说明：
1. 边距设置基于基准 DPI（150 DPI）
2. 导出时边距按 DPI 比例缩放
   - 150 DPI: 边距 10px = 10/150 = 0.067 英寸
   - 300 DPI: 边距 20px = 20/300 = 0.067 英寸（保持相同物理尺寸）
   - 600 DPI: 边距 40px = 40/600 = 0.067 英寸
3. PDF 页面尺寸根据裁剪后图像的物理尺寸设置
4. 结果：不同 DPI 的 PDF 页面尺寸完全一致，只有清晰度不同
```

### PDF 质量参数

**PDF 质量（1-100）** 控制 PDF 中嵌入图片的压缩程度：

- **100（无损）**: 不压缩，保留所有细节，文件最大
- **90-95**: 轻微压缩，肉眼几乎无损，文件适中（推荐）
- **70-85**: 中等压缩，适合网络分享
- **50以下**: 高压缩，文件小但质量明显下降

**注意**：当前版本使用 **SkiaSharp** 直接生成 PDF，质量参数暂未实现压缩控制，默认为高质量无损输出。

### DPI 选择建议

#### 1. 学术论文（推荐 300 DPI）
```
场景：提交给期刊或会议的论文配图
要求：高清晰度、可放大查看细节
推荐设置：
  - DPI: 300
  - PDF 质量: 100
  - 自动检测边界: 开启
  - 边距: 10-20 像素
```

#### 2. 日常文档（推荐 150 DPI）
```
场景：内部报告、演示文稿、教学材料
要求：清晰可读、文件大小适中
推荐设置：
  - DPI: 150（默认）
  - PDF 质量: 100
  - 自动检测边界: 开启
  - 边距: 10 像素
```

#### 3. 网页发布（推荐 72 DPI）
```
场景：网站、博客、在线分享
要求：快速加载、屏幕显示
推荐设置：
  - DPI: 72
  - PDF 质量: 85
  - 固定边距模式
```

#### 4. 专业印刷（推荐 600 DPI）
```
场景：书籍出版、海报印刷、展览
要求：极致清晰、可大幅放大
推荐设置：
  - DPI: 600
  - PDF 质量: 100
  - 精确边距控制
```

### 技术实现细节

FreeCut 使用以下技术栈生成高质量 PDF：

1. **图像导出**：PowerPoint COM 接口按指定 DPI 导出 PNG
2. **边距处理**：DPI 比例缩放算法确保物理尺寸一致
3. **PDF 生成**：SkiaSharp 库直接生成矢量 PDF
4. **尺寸计算**：
   ```csharp
   // 计算 PDF 页面物理尺寸（点，1英寸 = 72点）
   pageWidthInPoints = (imagePixelWidth / exportDPI) * 72
   pageHeightInPoints = (imagePixelHeight / exportDPI) * 72
   ```

### 常见问题

**Q: 为什么 600 DPI 的文件这么大？**
A: 600 DPI 的图像像素数是 150 DPI 的 16 倍（4× 宽度 × 4× 高度），因此文件大小会显著增加。只在真正需要高分辨率时使用。

**Q: 如何在保持清晰度的同时减小文件大小？**
A: 使用 300 DPI + 适当的 PDF 质量设置（85-95），这样可以在保持足够清晰度的同时大幅减小文件体积。

**Q: 边距设置会影响 PDF 页面大小吗？**
A: 是的。边距裁剪后，PDF 页面会完全匹配裁剪后图像的物理尺寸。使用自动检测边界 + 小边距可获得最紧凑的输出。



## 🎨 使用场景示例

### 📚 学术论文场景
```
任务：将PPT中的流程图导出到LaTeX论文中
操作步骤：
1. 在PowerPoint中选择包含流程图的幻灯片
2. 点击Ribbon中的"FreeCut设置"
3. 启用"自动检测边界"，设置边距为10像素
4. 选择300 DPI高质量导出
5. 点击"导出图片"，选择保存位置，得到PNG文件
6. 转换为PDF：
   - 使用在线工具将PNG转为PDF
   - 或在LaTeX中直接使用\includegraphics插入PNG
```

### 📊 研究报告场景
```
任务：批量导出图表制作报告附件
操作步骤：
1. 批量选择所有包含图表的幻灯片
2. 使用固定边距模式，设置统一的20px边距
3. 导出为高质量PNG图片，每张图表尺寸一致
4. 后续处理：
   - 使用图片编辑软件批量转换为PDF
   - 直接作为报告的图片附件使用
   - 或使用PDF合并工具创建单一PDF文档
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