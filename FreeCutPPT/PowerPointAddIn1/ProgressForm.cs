using System;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    /// <summary>
    /// 进度显示窗口
    /// </summary>
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            InitializeComponent();
            SetupForm();
            EnhanceFormAppearance();
        }

        private void EnhanceFormAppearance()
        {
            // 设置窗体背景色
            this.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);

            // 设置标题样式
            lblTitle.ForeColor = System.Drawing.Color.FromArgb(0, 102, 204);

            // 美化取消按钮
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.BackColor = System.Drawing.Color.FromArgb(128, 128, 128);
            btnCancel.ForeColor = System.Drawing.Color.White;
            btnCancel.FlatAppearance.BorderSize = 0;

            // 设置进度条样式（使用系统强调色）
            progressBar.ForeColor = System.Drawing.Color.FromArgb(0, 120, 212);
        }

        private void SetupForm()
        {
            // 设置窗口属性
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.ControlBox = false;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            // 设置进度条初始值
            progressBar.Value = 0;
            progressBar.Style = ProgressBarStyle.Continuous;

            // 设置初始文本
            lblStatus.Text = "准备中...";
            lblProgress.Text = "0%";
        }

        /// <summary>
        /// 设置进度值
        /// </summary>
        /// <param name="value">进度值 (0-100)</param>
        public void SetProgress(int value)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<int>(SetProgress), value);
                return;
            }

            try
            {
                // 确保进度值在有效范围内
                value = Math.Max(0, Math.Min(100, value));

                progressBar.Value = value;
                lblProgress.Text = $"{value}%";

                // 刷新界面
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"设置进度失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置状态文本
        /// </summary>
        /// <param name="text">状态文本</param>
        public void SetProgressText(string text)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(SetProgressText), text);
                return;
            }

            try
            {
                lblStatus.Text = text ?? "处理中...";

                // 刷新界面
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"设置状态文本失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置进度和状态文本
        /// </summary>
        /// <param name="value">进度值 (0-100)</param>
        /// <param name="text">状态文本</param>
        public void SetProgressAndText(int value, string text)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<int, string>(SetProgressAndText), value, text);
                return;
            }

            SetProgress(value);
            SetProgressText(text);
        }

        /// <summary>
        /// 显示错误状态
        /// </summary>
        /// <param name="errorMessage">错误消息</param>
        public void ShowError(string errorMessage)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(ShowError), errorMessage);
                return;
            }

            try
            {
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = 0;
                lblStatus.Text = $"错误: {errorMessage}";
                lblStatus.ForeColor = System.Drawing.Color.Red;
                lblProgress.Text = "失败";

                // 显示取消按钮
                btnCancel.Visible = true;
                btnCancel.Text = "关闭";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"显示错误状态失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 显示完成状态
        /// </summary>
        public void ShowCompleted()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(ShowCompleted));
                return;
            }

            try
            {
                progressBar.Value = 100;
                lblStatus.Text = "导出完成！";
                lblStatus.ForeColor = System.Drawing.Color.Green;
                lblProgress.Text = "100%";

                // 显示并启用关闭按钮
                btnCancel.Visible = true;
                btnCancel.Enabled = true;
                btnCancel.Text = "关闭";
                btnCancel.BringToFront();

                System.Diagnostics.Debug.WriteLine("ShowCompleted: 关闭按钮已显示并启用");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"显示完成状态失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 重置进度窗口
        /// </summary>
        public void Reset()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(Reset));
                return;
            }

            try
            {
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = 0;
                lblStatus.Text = "准备中...";
                lblStatus.ForeColor = System.Drawing.SystemColors.ControlText;
                lblProgress.Text = "0%";
                btnCancel.Visible = false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"重置进度窗口失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 设置为不确定模式（滚动条样式）
        /// </summary>
        /// <param name="text">状态文本</param>
        public void SetIndeterminate(string text = "处理中...")
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(SetIndeterminate), text);
                return;
            }

            try
            {
                progressBar.Style = ProgressBarStyle.Marquee;
                lblStatus.Text = text;
                lblProgress.Text = "进行中";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"设置不确定模式失败: {ex.Message}");
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("BtnCancel_Click: 按钮被点击");

            try
            {
                // 重置状态
                btnCancel.Visible = false;
                lblStatus.Text = "准备中...";
                lblStatus.ForeColor = System.Drawing.SystemColors.ControlText;
                progressBar.Value = 0;
                lblProgress.Text = "0%";

                System.Diagnostics.Debug.WriteLine("BtnCancel_Click: 准备隐藏窗口");
                this.Hide();
                System.Diagnostics.Debug.WriteLine("BtnCancel_Click: 窗口已隐藏");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"BtnCancel_Click 错误: {ex.Message}");
                MessageBox.Show($"关闭失败: {ex.Message}");
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 如果是用户点击关闭按钮，允许隐藏
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
            }
        }

        /// <summary>
        /// 设置详细信息文本
        /// </summary>
        /// <param name="details">详细信息</param>
        public void SetDetails(string details)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(SetDetails), details);
                return;
            }

            try
            {
                if (lblDetails != null)
                {
                    lblDetails.Text = details ?? "";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"设置详细信息失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 显示或隐藏取消按钮
        /// </summary>
        /// <param name="show">是否显示</param>
        public void ShowCancelButton(bool show)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<bool>(ShowCancelButton), show);
                return;
            }

            btnCancel.Visible = show;
        }
    }
}