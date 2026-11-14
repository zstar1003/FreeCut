# 安装后 PowerPoint 没有 FreeCut 标签页

本插件用于 **Windows 桌面版 Microsoft PowerPoint**。Mac、网页版 PowerPoint 和 WPS 不属于此 VSTO 插件的支持范围。

## 先检查安装包和运行环境

1. 完整解压 ZIP，不要在压缩包中直接运行安装器。安装器从自身所在目录查找插件，不依赖快捷方式的“起始位置”。
2. 安装 [.NET Framework 4.8](https://dotnet.microsoft.com/download/dotnet-framework/net48) 和 [Microsoft VSTO Runtime](https://aka.ms/vs/17/release/vs_vsto.exe)。未检测到 VSTO 时，新安装器会停止，不会继续显示安装成功。
3. 插件目录中必须同时包含同一次构建生成的非空文件：`PowerPointAddIn1.dll`、`PowerPointAddIn1.dll.manifest`、`PowerPointAddIn1.vsto`，以及构建输出的依赖 DLL。不要手写、修改或混用清单文件。新安装器会在修改安装目录/注册表前检查这三个文件是否存在且非空；这不等于验证签名、所有依赖或实际加载成功。
4. 关闭所有 PowerPoint 窗口，再安装。文件被占用或 `C:\FreeCut` 不可写时，安装器应显示具体错误；不要把错误当作成功。
5. 以实际使用 PowerPoint 的同一 Windows 用户运行安装器。注册表写入当前用户（HKCU），不是为所有用户安装。

## 确认文件和加载状态

在命令提示符运行以下只读检查：

```cmd
dir C:\FreeCut
reg query "HKCU\Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn"
```

应包含上述三个文件。`Manifest` 应指向 `C:\FreeCut\PowerPointAddIn1.vsto|vstolocal`；安装器设置 `LoadBehavior` 为 `3`，但 PowerPoint 加载失败后可能更改它，因此该值本身不是加载成功的证明。

重新打开 PowerPoint，进入“文件 → 选项 → 加载项”：

- 查看“COM 加载项 → 转到”中是否有 FreeCut、是否勾选，以及显示的加载错误。
- 查看“禁用项目”。若 PowerPoint 因崩溃禁用了插件，先记录错误原因再尝试重新启用，不要反复修改注册表强制加载。
- 若出现证书、信任或组织策略错误，请记录原文。不要关闭 Office 安全策略或删除整个 Office 缓存。

## 提交诊断信息

请在 issue 中提供 Windows 版本、PowerPoint 版本及 32/64 位、安装包版本、安装器状态文本、COM 加载项错误截图和上述命令结果（隐藏用户名等个人信息）。这些信息才能区分缺少运行时、包不完整、信任问题及 Office 禁用问题。

源码修复不会自动更新已有 release ZIP；使用旧安装包时，请先检查其版本。仓库没有 `FixFreeCut.bat` 自动修复脚本。
