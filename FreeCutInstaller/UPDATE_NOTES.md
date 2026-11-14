# FreeCut 安装器更新说明

## 安装前置检查修复

- 缺少 VSTO Runtime 时停止安装，不再允许忽略必需依赖后显示成功。
- 在复制和注册前检查 DLL、应用清单及部署清单存在且非空；复制后再次检查。
- 相对于安装器自身目录查找插件，避免从快捷方式启动时找不到文件。
- 安装器状态更新不再处理输入事件，避免安装/卸载过程被重复点击打断。
- “文件复制和注册完成”不代表已在 PowerPoint 中验证加载。Issue #1 尚无环境/错误详情，这些是确认的安装器缺陷修复，而非对该用户根因的确认。

## 🔧 早期修复说明

之前的安装器存在以下问题，导致插件显示"未加…1122 tokens truncated…f细节

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
