using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    /// <summary>
    /// 预览窗口
    /// </summary>
    public partial class PreviewForm : Form
    {
        private List<dynamic> slides;
        private CropSettings currentSettings;
        private int currentSlideIndex = 0;
        private List<Image> previewImages;

        public PreviewForm()
        {
            InitializeComponent();
            SetupForm();
            EnhanceFormAppearance();
            previewImages = new List<Image>();
        }

        /// <summary>
        /// 增强窗体外观
        /// </summary>
        private void EnhanceFormAppearance()
        {
            // 设置窗体背景色
            this.BackColor = System.Drawing.Color.FromArgb(240, 240, 240);

            // 设置标题样式
            lblTitle.ForeColor = System.Drawing.Color.FromArgb(0, 102, 204);

            // 美化按钮
            btnPrevious.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btnPrevious.BackColor = System.Drawing.Color.FromArgb(0, 120, 212);
            btnPrevious.ForeColor = System.Drawing.Color.White;
            btnPrevious.FlatAppearance.BorderSize = 0;

            btnNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btnNext.BackColor = System.Drawing.Color.FromArgb(0, 120, 212);
            btnNext.ForeColor = System.Drawing.Color.White;
            btnNext.FlatAppearance.BorderSize = 0;

            btnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btnRefresh.BackColor = System.Drawing.Color.FromArgb(0, 99, 177);
            btnRefresh.ForeColor = System.Drawing.Color.White;
            btnRefresh.FlatAppearance.BorderSize = 0;

            btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btnExport.BackColor = System.Drawing.Color.FromArgb(16, 124, 16);
            btnExport.ForeColor = System.Drawing.Color.White;
            btnExport.FlatAppearance.BorderSize = 0;

            btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            btnClose.BackColor = System.Drawing.Color.FromArgb(128, 128, 128);
            btnClose.ForeColor = System.Drawing.Color.White;
            btnClose.FlatAppearance.BorderSize = 0;

            // 设置导航面板背景
            panelNavigation.BackColor = System.Drawing.Color.White;
            panelPreview.BackColor = System.Drawing.Color.White;
        }

        private void SetupForm()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            this.KeyPreview = true;

            // 设置事件处理
            btnPrevious.Click += BtnPrevious_Click;
            btnNext.Click += BtnNext_Click;
            btnClose.Click += BtnClose_Click;
            btnExport.Click += BtnExport_Click;
            btnRefresh.Click += BtnRefresh_Click;

            // 键盘导航
            this.KeyDown += PreviewForm_KeyDown;

            // 设置初始状态
            UpdateNavigationButtons();
        }

        /// <summary>
        /// 加载预览
        /// </summary>
        /// <param name="slidesToPreview">要预览的幻灯片</param>
        public void LoadPreview(List<dynamic> slidesToPreview)
        {
            if (slidesToPreview == null || slidesToPreview.Count == 0)
            {
                MessageBox.Show("没有幻灯片可预览", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            this.slides = slidesToPreview;
            currentSlideIndex = 0;

            // 加载设置
            currentSettings = CropSettings.Load();

            // 生成预览图片
            GeneratePreviewImages();

            // 显示第一张预览
            ShowCurrentSlide();

            // 更新界面
            UpdateNavigationButtons();
            UpdateStatusText();
        }

        /// <summary>
        /// 生成预览图片
        /// </summary>
        private void GeneratePreviewImages()
        {
            try
            {
                // 清理之前的图片
                ClearPreviewImages();

                lblStatus.Text = "正在生成预览...";
                Application.DoEvents();

                for (int i = 0; i < slides.Count; i++)
                {
                    var slide = slides[i];
                    var previewImage = GenerateSlidePreview(slide, i + 1);
                    if (previewImage != null)
                    {
                        previewImages.Add(previewImage);
                    }
                }

                lblStatus.Text = $"预览生成完成，共 {previewImages.Count} 张";
            }
            catch (Exception ex)
            {
                lblStatus.Text = $"生成预览失败: {ex.Message}";
                MessageBox.Show($"生成预览失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 生成单张幻灯片预览
        /// </summary>
        private Image GenerateSlidePreview(dynamic slide, int slideNumber)
        {
            string tempImagePath = Path.Combine(Path.GetTempPath(), $"FreeCut_Preview_{slideNumber}_{Guid.NewGuid()}.png");

            try
            {
                // 导出幻灯片为临时图片（较低DPI以提高速度）
                slide.Export(tempImagePath, "PNG", 150, 150);

                // 加载并处理图片
                using (var originalImage = Image.FromFile(tempImagePath))
                {
                    var exporter = new PdfExporter(currentSettings);

                    // 使用反射调用私有方法进行裁剪（这里简化处理）
                    var croppedImage = CropImageForPreview(originalImage);

                    return new Bitmap(croppedImage);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"生成幻灯片 {slideNumber} 预览失败: {ex.Message}");
                return null;
            }
            finally
            {
                // 清理临时文件
                if (File.Exists(tempImagePath))
                {
                    try { File.Delete(tempImagePath); }
                    catch { }
                }
            }
        }

        /// <summary>
        /// 为预览裁剪图片（简化版本）
        /// </summary>
        private Image CropImageForPreview(Image originalImage)
        {
            if (!currentSettings.AutoDetectBounds)
            {
                // 固定边距裁剪
                var cropRect = new Rectangle(
                    currentSettings.LeftMargin,
                    currentSettings.TopMargin,
                    Math.Max(1, originalImage.Width - currentSettings.LeftMargin - currentSettings.RightMargin),
                    Math.Max(1, originalImage.Height - currentSettings.TopMargin - currentSettings.BottomMargin)
                );

                return CropBitmap(originalImage, cropRect);
            }
            else
            {
                // 自动检测（简化版本）
                using (var bitmap = new Bitmap(originalImage))
                {
                    var contentBounds = DetectContentBoundsSimple(bitmap);

                    var cropRect = new Rectangle(
                        Math.Max(0, contentBounds.Left - currentSettings.LeftMargin),
                        Math.Max(0, contentBounds.Top - currentSettings.TopMargin),
                        Math.Min(bitmap.Width - Math.Max(0, contentBounds.Left - currentSettings.LeftMargin),
                                contentBounds.Width + currentSettings.LeftMargin + currentSettings.RightMargin),
                        Math.Min(bitmap.Height - Math.Max(0, contentBounds.Top - currentSettings.TopMargin),
                                contentBounds.Height + currentSettings.TopMargin + currentSettings.BottomMargin)
                    );

                    return CropBitmap(bitmap, cropRect);
                }
            }
        }

        /// <summary>
        /// 简化的内容边界检测
        /// </summary>
        private Rectangle DetectContentBoundsSimple(Bitmap bitmap)
        {
            // 简化实现：使用左上角像素作为背景色
            Color backgroundColor = bitmap.GetPixel(0, 0);
            int tolerance = currentSettings.DetectionTolerance;

            int left = bitmap.Width, right = 0, top = bitmap.Height, bottom = 0;

            // 采样检测以提高速度
            int step = Math.Max(1, bitmap.Width / 100);

            for (int y = 0; y < bitmap.Height; y += step)
            {
                for (int x = 0; x < bitmap.Width; x += step)
                {
                    Color pixelColor = bitmap.GetPixel(x, y);
                    if (!IsBackgroundColorSimple(pixelColor, backgroundColor, tolerance))
                    {
                        left = Math.Min(left, x);
                        right = Math.Max(right, x);
                        top = Math.Min(top, y);
                        bottom = Math.Max(bottom, y);
                    }
                }
            }

            if (left >= right || top >= bottom)
            {
                return new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            }

            return new Rectangle(left, top, right - left + 1, bottom - top + 1);
        }

        private bool IsBackgroundColorSimple(Color pixelColor, Color backgroundColor, int tolerance)
        {
            int rDiff = Math.Abs(pixelColor.R - backgroundColor.R);
            int gDiff = Math.Abs(pixelColor.G - backgroundColor.G);
            int bDiff = Math.Abs(pixelColor.B - backgroundColor.B);

            return rDiff <= tolerance && gDiff <= tolerance && bDiff <= tolerance;
        }

        private Image CropBitmap(Image source, Rectangle cropRect)
        {
            cropRect = Rectangle.Intersect(cropRect, new Rectangle(0, 0, source.Width, source.Height));

            if (cropRect.IsEmpty)
            {
                return new Bitmap(source);
            }

            var croppedBitmap = new Bitmap(cropRect.Width, cropRect.Height);
            using (var graphics = Graphics.FromImage(croppedBitmap))
            {
                // 设置高质量渲染模式
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                graphics.DrawImage(source, 0, 0, cropRect, GraphicsUnit.Pixel);
            }

            return croppedBitmap;
        }

        /// <summary>
        /// 显示当前幻灯片
        /// </summary>
        private void ShowCurrentSlide()
        {
            if (previewImages.Count == 0 || currentSlideIndex < 0 || currentSlideIndex >= previewImages.Count)
            {
                pictureBoxPreview.Image = null;
                return;
            }

            pictureBoxPreview.Image = previewImages[currentSlideIndex];
            UpdateStatusText();
        }

        /// <summary>
        /// 更新状态文本
        /// </summary>
        private void UpdateStatusText()
        {
            if (slides != null && slides.Count > 0)
            {
                lblSlideInfo.Text = $"第 {currentSlideIndex + 1} / {slides.Count} 张幻灯片";
            }
            else
            {
                lblSlideInfo.Text = "无幻灯片";
            }
        }

        /// <summary>
        /// 更新导航按钮状态
        /// </summary>
        private void UpdateNavigationButtons()
        {
            btnPrevious.Enabled = slides != null && currentSlideIndex > 0;
            btnNext.Enabled = slides != null && currentSlideIndex < slides.Count - 1;
            btnExport.Enabled = slides != null && slides.Count > 0;
        }

        /// <summary>
        /// 清理预览图片
        /// </summary>
        private void ClearPreviewImages()
        {
            foreach (var image in previewImages)
            {
                image?.Dispose();
            }
            previewImages.Clear();
        }

        private void BtnPrevious_Click(object sender, EventArgs e)
        {
            if (currentSlideIndex > 0)
            {
                currentSlideIndex--;
                ShowCurrentSlide();
                UpdateNavigationButtons();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            if (slides != null && currentSlideIndex < slides.Count - 1)
            {
                currentSlideIndex++;
                ShowCurrentSlide();
                UpdateNavigationButtons();
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取当前预览的幻灯片
                if (slides == null || slides.Count == 0)
                {
                    MessageBox.Show("没有可导出的幻灯片", "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 获取保存路径
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "PDF文件|*.pdf";
                    saveDialog.Title = "选择PDF保存位置";
                    saveDialog.FileName = $"PPT导出_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        // 执行导出
                        var settings = CropSettings.Load();
                        var progressForm = new ProgressForm();
                        progressForm.Show();
                        progressForm.SetProgressText("准备导出...");
                        progressForm.SetProgress(0);

                        var exporter = new PdfExporter(settings);
                        exporter.ExportToPdf(slides, saveDialog.FileName, (progress, message) =>
                        {
                            progressForm.SetProgress(progress);
                            progressForm.SetProgressText(message);
                        });

                        progressForm.SetProgress(100);
                        progressForm.SetProgressText("导出完成！");
                        progressForm.ShowCompleted();
                    }
                }

                // 关闭预览窗口
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            if (slides != null)
            {
                // 重新加载设置
                currentSettings = CropSettings.Load();

                // 重新生成预览
                GeneratePreviewImages();
                ShowCurrentSlide();
            }
        }

        private void PreviewForm_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Left:
                case Keys.Up:
                    BtnPrevious_Click(sender, EventArgs.Empty);
                    e.Handled = true;
                    break;
                case Keys.Right:
                case Keys.Down:
                    BtnNext_Click(sender, EventArgs.Empty);
                    e.Handled = true;
                    break;
                case Keys.Escape:
                    BtnClose_Click(sender, EventArgs.Empty);
                    e.Handled = true;
                    break;
                case Keys.F5:
                    BtnRefresh_Click(sender, EventArgs.Empty);
                    e.Handled = true;
                    break;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 隐藏而不是关闭窗口
            e.Cancel = true;
            this.Hide();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                ClearPreviewImages();
                components?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}