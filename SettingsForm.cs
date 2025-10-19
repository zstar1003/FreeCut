using System;
using System.Windows.Forms;

namespace FreeCut
{
    /// <summary>
    /// 设置窗口
    /// </summary>
    public partial class SettingsForm : Form
    {
        private CropSettings currentSettings;

        public SettingsForm()
        {
            InitializeComponent();
            LoadSettings();
            SetupEventHandlers();
        }

        private void LoadSettings()
        {
            try
            {
                currentSettings = CropSettings.Load();
                ApplySettingsToControls();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载设置失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                currentSettings = new CropSettings();
            }
        }

        private void ApplySettingsToControls()
        {
            // 边距设置
            numTopMargin.Value = currentSettings.TopMargin;
            numBottomMargin.Value = currentSettings.BottomMargin;
            numLeftMargin.Value = currentSettings.LeftMargin;
            numRightMargin.Value = currentSettings.RightMargin;

            // 自动检测设置
            chkAutoDetect.Checked = currentSettings.AutoDetectBounds;
            numTolerance.Value = currentSettings.DetectionTolerance;

            // 导出设置
            numPdfQuality.Value = currentSettings.PdfQuality;
            chkPreserveAspectRatio.Checked = currentSettings.PreserveAspectRatio;

            // DPI设置
            cmbDpi.Items.Clear();
            var dpis = CropSettings.GetAvailableDpis();
            foreach (var dpi in dpis)
            {
                cmbDpi.Items.Add(CropSettings.GetDpiDescription(dpi));
            }

            // 设置当前DPI
            for (int i = 0; i < dpis.Length; i++)
            {
                if (dpis[i] == currentSettings.ExportDpi)
                {
                    cmbDpi.SelectedIndex = i;
                    break;
                }
            }

            // 背景模式
            cmbBackgroundMode.Items.Clear();
            cmbBackgroundMode.Items.Add("自动检测");
            cmbBackgroundMode.Items.Add("白色背景");
            cmbBackgroundMode.Items.Add("透明背景");
            cmbBackgroundMode.Items.Add("自定义颜色");
            cmbBackgroundMode.SelectedIndex = (int)currentSettings.BackgroundMode;

            UpdateControlStates();
        }

        private void SetupEventHandlers()
        {
            // 边距控件事件
            numTopMargin.ValueChanged += OnMarginValueChanged;
            numBottomMargin.ValueChanged += OnMarginValueChanged;
            numLeftMargin.ValueChanged += OnMarginValueChanged;
            numRightMargin.ValueChanged += OnMarginValueChanged;

            // 自动检测事件
            chkAutoDetect.CheckedChanged += AutoDetectChanged;
            numTolerance.ValueChanged += (s, e) => ValidateAndUpdateSettings();

            // 导出设置事件
            numPdfQuality.ValueChanged += (s, e) => ValidateAndUpdateSettings();
            chkPreserveAspectRatio.CheckedChanged += (s, e) => ValidateAndUpdateSettings();
            cmbDpi.SelectedIndexChanged += (s, e) => ValidateAndUpdateSettings();
            cmbBackgroundMode.SelectedIndexChanged += BackgroundModeChanged;

            // 按钮事件
            btnSave.Click += BtnSave_Click;
            btnReset.Click += BtnReset_Click;
            btnCancel.Click += BtnCancel_Click;
            btnPreview.Click += BtnPreview_Click;
            btnExport.Click += BtnExport_Click;
            btnSetAllMargins.Click += BtnSetAllMargins_Click;
        }

        private void OnMarginValueChanged(object sender, EventArgs e)
        {
            // 更新对称边距按钮状态
            bool allSame = numTopMargin.Value == numBottomMargin.Value &&
                          numBottomMargin.Value == numLeftMargin.Value &&
                          numLeftMargin.Value == numRightMargin.Value;

            btnSetAllMargins.Text = allSame ? "边距已统一" : "统一边距";
            btnSetAllMargins.Enabled = !allSame;

            ValidateAndUpdateSettings();
        }

        private void AutoDetectChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
            ValidateAndUpdateSettings();
        }

        private void BackgroundModeChanged(object sender, EventArgs e)
        {
            UpdateControlStates();
            ValidateAndUpdateSettings();
        }

        private void UpdateControlStates()
        {
            // 容差控件状态
            numTolerance.Enabled = chkAutoDetect.Checked;
            lblTolerance.Enabled = chkAutoDetect.Checked;

            // 背景色控件状态
            bool customColor = cmbBackgroundMode.SelectedIndex == 3; // Custom
            btnCustomColor.Enabled = customColor;
            panelColorPreview.Enabled = customColor;

            // 边距提示
            if (chkAutoDetect.Checked)
            {
                lblMarginHint.Text = "边距将作为检测到的内容边界的额外缓冲区";
            }
            else
            {
                lblMarginHint.Text = "边距将从图片边缘开始计算";
            }
        }

        private void ValidateAndUpdateSettings()
        {
            try
            {
                // 更新设置对象
                currentSettings.TopMargin = (int)numTopMargin.Value;
                currentSettings.BottomMargin = (int)numBottomMargin.Value;
                currentSettings.LeftMargin = (int)numLeftMargin.Value;
                currentSettings.RightMargin = (int)numRightMargin.Value;
                currentSettings.AutoDetectBounds = chkAutoDetect.Checked;
                currentSettings.DetectionTolerance = (int)numTolerance.Value;
                currentSettings.PdfQuality = (int)numPdfQuality.Value;
                currentSettings.PreserveAspectRatio = chkPreserveAspectRatio.Checked;

                if (cmbDpi.SelectedIndex >= 0)
                {
                    var dpis = CropSettings.GetAvailableDpis();
                    currentSettings.ExportDpi = dpis[cmbDpi.SelectedIndex];
                }

                if (cmbBackgroundMode.SelectedIndex >= 0)
                {
                    currentSettings.BackgroundMode = (BackgroundDetectionMode)cmbBackgroundMode.SelectedIndex;
                }

                // 验证设置
                string errorMessage;
                bool isValid = currentSettings.Validate(out errorMessage);

                btnSave.Enabled = isValid;
                btnExport.Enabled = isValid;

                if (!isValid)
                {
                    lblValidation.Text = errorMessage;
                    lblValidation.ForeColor = System.Drawing.Color.Red;
                }
                else
                {
                    lblValidation.Text = "设置有效";
                    lblValidation.ForeColor = System.Drawing.Color.Green;
                }
            }
            catch (Exception ex)
            {
                lblValidation.Text = $"验证失败：{ex.Message}";
                lblValidation.ForeColor = System.Drawing.Color.Red;
                btnSave.Enabled = false;
                btnExport.Enabled = false;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                currentSettings.Save();
                MessageBox.Show("设置已保存", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("确定要重置所有设置为默认值吗？", "确认重置",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                currentSettings.ResetToDefault();
                ApplySettingsToControls();
                MessageBox.Show("设置已重置为默认值", "重置完成",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void BtnPreview_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取当前选中的幻灯片
                var selectedSlides = ThisAddIn.GetSelectedSlides();
                if (selectedSlides.Count == 0)
                {
                    MessageBox.Show("请先在PowerPoint中选择要预览的幻灯片", "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 显示预览窗口
                var ribbon = new FreeCutRibbon();
                ribbon.ShowPreview();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"预览失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            try
            {
                // 先保存当前设置
                currentSettings.Save();

                // 触发导出
                var ribbon = new FreeCutRibbon();
                ribbon.OnExportPDF(null);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSetAllMargins_Click(object sender, EventArgs e)
        {
            // 将所有边距设置为上边距的值
            decimal marginValue = numTopMargin.Value;
            numBottomMargin.Value = marginValue;
            numLeftMargin.Value = marginValue;
            numRightMargin.Value = marginValue;
        }

        private void BtnCustomColor_Click(object sender, EventArgs e)
        {
            using (var colorDialog = new ColorDialog())
            {
                colorDialog.Color = System.Drawing.Color.FromArgb((int)currentSettings.CustomBackgroundColor);

                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    currentSettings.CustomBackgroundColor = (uint)colorDialog.Color.ToArgb();
                    panelColorPreview.BackColor = colorDialog.Color;
                    ValidateAndUpdateSettings();
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 隐藏而不是关闭窗口
            e.Cancel = true;
            this.Hide();
        }
    }
}