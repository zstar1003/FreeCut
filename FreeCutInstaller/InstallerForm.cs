using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Microsoft.Win32;

namespace FreeCutInstaller
{
    public partial class InstallerForm : Form
    {
        private const string INSTALL_PATH = @"C:\FreeCut";

        private TextBox statusTextBox;

        public InstallerForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Form
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(500, 400);
            this.Text = "FreeCut PowerPoint 插件安装器";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Title Label
            var titleLabel = new Label();
            titleLabel.Text = "FreeCut - PPT自动裁剪PDF导出插件";
            titleLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold);
            titleLabel.Size = new System.Drawing.Size(460, 30);
            titleLabel.Location = new System.Drawing.Point(20, 20);
            titleLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;

            // Description Label
            var descLabel = new Label();
            descLabel.Text = "这个安装器将为您的PowerPoint安装FreeCut插件。\n\n功能特性：\n• 智能页面选择和自动边界检测\n• 自定义边距设置 (0-500像素)\n• 高质量PDF导出 (72-600 DPI)\n• 实时预览和批量处理\n• 设置持久化存储";
            descLabel.Size = new System.Drawing.Size(460, 120);
            descLabel.Location = new System.Drawing.Point(20, 60);

            // Status TextBox
            statusTextBox = new TextBox();
            statusTextBox.Name = "statusTextBox";
            statusTextBox.Multiline = true;
            statusTextBox.ScrollBars = ScrollBars.Vertical;
            statusTextBox.ReadOnly = true;
            statusTextBox.Size = new System.Drawing.Size(460, 120);
            statusTextBox.Location = new System.Drawing.Point(20, 190);
            statusTextBox.Text = "准备安装...\r\n";

            // Install Button
            var installButton = new Button();
            installButton.Text = "安装插件";
            installButton.Size = new System.Drawing.Size(100, 35);
            installButton.Location = new System.Drawing.Point(150, 330);
            installButton.UseVisualStyleBackColor = true;
            installButton.Click += InstallButton_Click;

            // Uninstall Button
            var uninstallButton = new Button();
            uninstallButton.Text = "卸载插件";
            uninstallButton.Size = new System.Drawing.Size(100, 35);
            uninstallButton.Location = new System.Drawing.Point(260, 330);
            uninstallButton.UseVisualStyleBackColor = true;
            uninstallButton.Click += UninstallButton_Click;

            // Close Button
            var closeButton = new Button();
            closeButton.Text = "关闭";
            closeButton.Size = new System.Drawing.Size(80, 35);
            closeButton.Location = new System.Drawing.Point(370, 330);
            closeButton.UseVisualStyleBackColor = true;
            closeButton.Click += (s, e) => this.Close();

            // Add controls to form
            this.Controls.Add(titleLabel);
            this.Controls.Add(descLabel);
            this.Controls.Add(statusTextBox);
            this.Controls.Add(installButton);
            this.Controls.Add(uninstallButton);
            this.Controls.Add(closeButton);

            this.ResumeLayout(false);
        }

        private void InstallButton_Click(object sender, EventArgs e)
        {
            try
            {
                statusTextBox.Clear();
                LogMessage("开始安装FreeCut插件...");

                // 检查 VSTO Runtime
                LogMessage("检查 VSTO Runtime...");
                if (!CheckVSTORuntime())
                {
                    LogMessage("安装已停止: 未检测到 VSTO Runtime，尚未复制或注册插件。");
                    MessageBox.Show(
                        "未检测到 Microsoft Visual Studio Tools for Office Runtime。\n\n" +
                        "请先安装该组件，再重新运行安装器：\n" +
                        "https://aka.ms/vs/17/release/vs_vsto.exe",
                        "缺少必需组件", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                else
                {
                    LogMessage("VSTO Runtime 已安装");
                }

                // 必须在修改安装目录/注册表之前检查完整的部署包。
                string sourceDir = InstallPackage.FindSourceDirectory(AppDomain.CurrentDomain.BaseDirectory);
                if (sourceDir == null)
                {
                    throw new FileNotFoundException("找不到 PowerPointAddIn1.dll，请完整解压安装包后再运行安装器。");
                }
                InstallPackage.Validate(sourceDir);
                LogMessage("找到完整插件包: " + sourceDir);

                // 创建安装目录
                LogMessage("创建安装目录...");
                if (!Directory.Exists(INSTALL_PATH))
                {
                    Directory.CreateDirectory(INSTALL_PATH);
                }

                // 复制文件
                LogMessage("复制插件文件...");

                // 复制所有 DLL 文件
                foreach (string file in Directory.GetFiles(sourceDir, "*.dll"))
                {
                    string fileName = Path.GetFileName(file);
                    string destFile = Path.Combine(INSTALL_PATH, fileName);
                    InstallPackage.CopyFile(file, destFile);
                    LogMessage("已复制: " + fileName);
                }

                // 复制 .vsto 和 .manifest 文件（VSTO 部署必需）
                foreach (string file in Directory.GetFiles(sourceDir, "*.vsto"))
                {
                    string fileName = Path.GetFileName(file);
                    string destFile = Path.Combine(INSTALL_PATH, fileName);
                    InstallPackage.CopyFile(file, destFile);
                    LogMessage("已复制: " + fileName);
                }

                foreach (string file in Directory.GetFiles(sourceDir, "*.manifest"))
                {
                    string fileName = Path.GetFileName(file);
                    string destFile = Path.Combine(INSTALL_PATH, fileName);
                    InstallPackage.CopyFile(file, destFile);
                    LogMessage("已复制: " + fileName);
                }

                // 复制PDB文件（如果存在，用于调试）
                foreach (string file in Directory.GetFiles(sourceDir, "*.pdb"))
                {
                    string fileName = Path.GetFileName(file);
                    string destFile = Path.Combine(INSTALL_PATH, fileName);
                    InstallPackage.CopyFile(file, destFile);
                    LogMessage("已复制: " + fileName);
                }

                // 注册插件
                LogMessage("注册插件到PowerPoint...");
                InstallPackage.Validate(INSTALL_PATH);
                RegisterPlugin();

                LogMessage("");
                LogMessage("文件复制和注册完成；加载结果请在 PowerPoint 中确认。");
                LogMessage("");
                LogMessage("请重启PowerPoint，然后查看Ribbon中的FreeCut标签页。");

                string successMessage = "FreeCut 插件文件复制和注册完成。\n\n请重启 PowerPoint 查看 FreeCut 标签页。\n若未出现，请参阅安装包的 TROUBLESHOOTING.md。";

                MessageBox.Show(successMessage,
                    "安装成功", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 启动PowerPoint
                if (MessageBox.Show("是否现在启动PowerPoint进行测试？", "启动PowerPoint",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        Process.Start("powerpnt.exe");
                    }
                    catch
                    {
                        MessageBox.Show("无法自动启动PowerPoint，请手动启动。", "提示",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage("安装失败: " + ex.Message);
                MessageBox.Show("安装失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UninstallButton_Click(object sender, EventArgs e)
        {
            try
            {
                statusTextBox.Clear();
                LogMessage("开始卸载FreeCut插件...");

                // 删除注册表项
                LogMessage("清理注册表...");
                UnregisterPlugin();

                // 删除文件
                LogMessage("删除插件文件...");
                if (Directory.Exists(INSTALL_PATH))
                {
                    Directory.Delete(INSTALL_PATH, true);
                    LogMessage("插件文件已删除");
                }

                LogMessage("");
                LogMessage("卸载完成！");
                LogMessage("请重启PowerPoint以确保插件完全卸载。");

                MessageBox.Show("FreeCut插件已成功卸载！\n\n请重启PowerPoint。",
                    "卸载成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LogMessage("卸载失败: " + ex.Message);
                MessageBox.Show("卸载失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RegisterPlugin()
        {
            // VSTO 插件使用 .vsto 文件进行部署
            string vstoPath = Path.Combine(INSTALL_PATH, "PowerPointAddIn1.vsto");

            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn"))
            {
                key.SetValue("Description", "FreeCut - PPT自动裁剪PDF导出插件", RegistryValueKind.String);
                key.SetValue("FriendlyName", "FreeCut", RegistryValueKind.String);
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                key.SetValue("Manifest", vstoPath + "|vstolocal", RegistryValueKind.String);
            }
        }

        private bool CheckVSTORuntime()
        {
            // 检查 VSTO Runtime 是否安装
            // VSTO 2010 Runtime (v10.0) 注册表路径
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"))
                {
                    if (key != null)
                    {
                        return true;
                    }
                }

                // 也检查 32 位注册表路径（在 64 位系统上）
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\VSTO Runtime Setup\v4R"))
                {
                    if (key != null)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                // 忽略错误
            }

            return false;
        }

        private void UnregisterPlugin()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKey(@"Software\Microsoft\Office\PowerPoint\Addins\PowerPointAddIn1.ThisAddIn");
            }
            catch
            {
                // 忽略删除失败的错误
            }
        }

        private void LogMessage(string message)
        {
            statusTextBox.AppendText(message + "\r\n");
            statusTextBox.SelectionStart = statusTextBox.Text.Length;
            statusTextBox.ScrollToCaret();
            // Do not pump input here: install/uninstall must not re-enter mid-copy.
            statusTextBox.Refresh();
        }
    }
}
