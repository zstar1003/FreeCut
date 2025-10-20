# FreeCut 快速开始指南

## 🚀 3 分钟快速安装

### 第一步：构建项目

1. 打开 Visual Studio
2. 打开项目文件：`FreeCutPPT\PowerPointAddIn1\PowerPointAddIn1.sln`
3. 顶部工具栏选择 **Release** 配置
4. 按 `Ctrl+Shift+B` 或点击 `生成` → `重新生成解决方案`
5. 等待构建完成（应该显示"生成成功"）

### 第二步：安装插件

**Windows 用户（推荐）：**
- 双击运行 `Install-FreeCut.bat`
- 输入 `1` 并按回车
- 等待安装完成

**或者使用 PowerShell：**
- 右键点击 `Install-FreeCut.ps1`
- 选择"使用 PowerShell 运行"
- 输入 `1` 并按回车

### 第三步：使用插件

1. 重启 PowerPoint
2. 在 Ribbon 顶部找到 **FreeCut** 标签页
3. 选择一个或多个幻灯片
4. 点击 **FreeCut设置** 调整参数
5. 点击 **导出PDF** 保存文件

---

## 📖 完整文档

- **安装指南**: [INSTALL.md](INSTALL.md) - 详细安装步骤和故障排除
- **项目说明**: [README.md](README.md) - 功能介绍和技术说明

---

## ⚡ 快捷操作

### 开发调试
在 Visual Studio 中按 `F5`，会自动启动 PowerPoint 并加载插件。

### 卸载插件
运行安装脚本，选择 `2. 卸载插件`

### 推荐设置

**学术论文（高质量）：**
- DPI: 300
- PDF 质量: 100
- 自动检测边界: 开启
- 边距: 10-20 像素

**日常文档（平衡）：**
- DPI: 150（默认）
- PDF 质量: 100
- 自动检测边界: 开启
- 边距: 10 像素

**网页发布（小文件）：**
- DPI: 72
- PDF 质量: 85
- 固定边距模式

---

## ❓ 遇到问题？

### 安装后看不到 FreeCut 标签页

1. **检查加载项**
   - PowerPoint → `文件` → `选项` → `加载项`
   - 选择 `COM 加载项` → `转到`
   - 确保 `FreeCut` 已勾选

2. **清除缓存**
   - 关闭所有 PowerPoint
   - 删除文件夹：`%localappdata%\Apps\2.0`
   - 重新运行安装脚本

### 找不到 DLL 文件

确保在 Visual Studio 中完成了项目构建（第一步）。

### PowerShell 脚本无法运行

使用 `Install-FreeCut.bat` 批处理脚本即可，无需 PowerShell。

---

## 💡 主要功能

✅ 自动裁剪 PPT 边界，无需手动调整
✅ 支持 72/150/300/600 DPI 高质量导出
✅ 批量处理多页幻灯片
✅ 自定义边距控制（0-500 像素）
✅ DPI 自适应缩放，保持物理尺寸一致

---

**🎉 开始使用 FreeCut，让 PPT 导出更轻松！🎉**
