using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace FreeCut
{
    /// <summary>
    /// PDF导出器
    /// </summary>
    public class PdfExporter
    {
        private readonly CropSettings settings;

        public PdfExporter(CropSettings settings)
        {
            this.settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <summary>
        /// 导出幻灯片为PDF
        /// </summary>
        public void ExportToPdf(List<dynamic> slides, string outputPath, Action<int, string> progressCallback = null)
        {
            if (slides == null || slides.Count == 0)
                throw new ArgumentException("幻灯片列表不能为空", nameof(slides));

            progressCallback?.Invoke(0, "初始化PDF文档...");

            using (var document = new Document())
            {
                using (var fileStream = new FileStream(outputPath, FileMode.Create))
                {
                    var writer = PdfWriter.GetInstance(document, fileStream);
                    document.Open();

                    for (int i = 0; i < slides.Count; i++)
                    {
                        var slide = slides[i];
                        var progress = (int)((float)(i + 1) / slides.Count * 100);
                        progressCallback?.Invoke(progress, $"正在处理第 {i + 1}/{slides.Count} 张幻灯片...");

                        try
                        {
                            ProcessSlide(document, slide, i + 1);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"处理第 {i + 1} 张幻灯片失败: {ex.Message}");
                            // 继续处理其他幻灯片
                        }
                    }

                    document.Close();
                }
            }

            progressCallback?.Invoke(100, "PDF导出完成");
        }

        /// <summary>
        /// 处理单张幻灯片
        /// </summary>
        private void ProcessSlide(Document document, dynamic slide, int slideNumber)
        {
            // 导出幻灯片为图片
            string tempImagePath = Path.Combine(Path.GetTempPath(), $"FreeCut_Slide_{slideNumber}_{Guid.NewGuid()}.png");

            try
            {
                // 使用PowerPoint导出幻灯片为图片
                slide.Export(tempImagePath, "PNG", settings.ExportDpi, settings.ExportDpi);

                // 处理图片
                using (var originalImage = System.Drawing.Image.FromFile(tempImagePath))
                {
                    var croppedImage = CropImage(originalImage);

                    // 将处理后的图片添加到PDF
                    AddImageToPdf(document, croppedImage);
                }
            }
            finally
            {
                // 清理临时文件
                if (File.Exists(tempImagePath))
                {
                    File.Delete(tempImagePath);
                }
            }
        }

        /// <summary>
        /// 裁剪图片
        /// </summary>
        private System.Drawing.Image CropImage(System.Drawing.Image originalImage)
        {
            if (settings.AutoDetectBounds)
            {
                return CropImageWithAutoDetection(originalImage);
            }
            else
            {
                return CropImageWithFixedMargins(originalImage);
            }
        }

        /// <summary>
        /// 自动检测边界并裁剪
        /// </summary>
        private System.Drawing.Image CropImageWithAutoDetection(System.Drawing.Image originalImage)
        {
            using (var bitmap = new Bitmap(originalImage))
            {
                // 检测内容边界
                var contentBounds = DetectContentBounds(bitmap);

                // 应用边距
                var cropRect = new System.Drawing.Rectangle(
                    Math.Max(0, contentBounds.Left - settings.LeftMargin),
                    Math.Max(0, contentBounds.Top - settings.TopMargin),
                    Math.Min(bitmap.Width - Math.Max(0, contentBounds.Left - settings.LeftMargin),
                            contentBounds.Width + settings.LeftMargin + settings.RightMargin),
                    Math.Min(bitmap.Height - Math.Max(0, contentBounds.Top - settings.TopMargin),
                            contentBounds.Height + settings.TopMargin + settings.BottomMargin)
                );

                return CropBitmap(bitmap, cropRect);
            }
        }

        /// <summary>
        /// 使用固定边距裁剪
        /// </summary>
        private System.Drawing.Image CropImageWithFixedMargins(System.Drawing.Image originalImage)
        {
            var cropRect = new System.Drawing.Rectangle(
                settings.LeftMargin,
                settings.TopMargin,
                originalImage.Width - settings.LeftMargin - settings.RightMargin,
                originalImage.Height - settings.TopMargin - settings.BottomMargin
            );

            using (var bitmap = new Bitmap(originalImage))
            {
                return CropBitmap(bitmap, cropRect);
            }
        }

        /// <summary>
        /// 检测内容边界
        /// </summary>
        private System.Drawing.Rectangle DetectContentBounds(Bitmap bitmap)
        {
            // 获取背景色
            Color backgroundColor = GetBackgroundColor(bitmap);

            int left = bitmap.Width;
            int right = 0;
            int top = bitmap.Height;
            int bottom = 0;

            // 扫描整个图片找到内容边界
            for (int y = 0; y < bitmap.Height; y++)
            {
                for (int x = 0; x < bitmap.Width; x++)
                {
                    Color pixelColor = bitmap.GetPixel(x, y);

                    if (!IsBackgroundColor(pixelColor, backgroundColor, settings.DetectionTolerance))
                    {
                        // 找到内容像素
                        left = Math.Min(left, x);
                        right = Math.Max(right, x);
                        top = Math.Min(top, y);
                        bottom = Math.Max(bottom, y);
                    }
                }
            }

            // 如果没有找到内容，返回整个图片
            if (left >= right || top >= bottom)
            {
                return new System.Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);
            }

            return new System.Drawing.Rectangle(left, top, right - left + 1, bottom - top + 1);
        }

        /// <summary>
        /// 获取背景色
        /// </summary>
        private Color GetBackgroundColor(Bitmap bitmap)
        {
            switch (settings.BackgroundMode)
            {
                case BackgroundDetectionMode.White:
                    return Color.White;

                case BackgroundDetectionMode.Transparent:
                    return Color.Transparent;

                case BackgroundDetectionMode.Custom:
                    return Color.FromArgb((int)settings.CustomBackgroundColor);

                case BackgroundDetectionMode.AutoDetect:
                default:
                    // 使用图片四个角的最常见颜色作为背景色
                    var cornerColors = new[]
                    {
                        bitmap.GetPixel(0, 0),
                        bitmap.GetPixel(bitmap.Width - 1, 0),
                        bitmap.GetPixel(0, bitmap.Height - 1),
                        bitmap.GetPixel(bitmap.Width - 1, bitmap.Height - 1)
                    };

                    // 简单实现：使用左上角颜色
                    return cornerColors[0];
            }
        }

        /// <summary>
        /// 判断是否为背景色
        /// </summary>
        private bool IsBackgroundColor(Color pixelColor, Color backgroundColor, int tolerance)
        {
            if (backgroundColor == Color.Transparent)
            {
                return pixelColor.A < 128; // 半透明以上认为是内容
            }

            // 计算颜色差异
            int rDiff = Math.Abs(pixelColor.R - backgroundColor.R);
            int gDiff = Math.Abs(pixelColor.G - backgroundColor.G);
            int bDiff = Math.Abs(pixelColor.B - backgroundColor.B);

            return rDiff <= tolerance && gDiff <= tolerance && bDiff <= tolerance;
        }

        /// <summary>
        /// 裁剪位图
        /// </summary>
        private System.Drawing.Image CropBitmap(Bitmap source, System.Drawing.Rectangle cropRect)
        {
            // 确保裁剪区域在图片范围内
            cropRect = System.Drawing.Rectangle.Intersect(cropRect,
                new System.Drawing.Rectangle(0, 0, source.Width, source.Height));

            if (cropRect.IsEmpty)
            {
                return new Bitmap(source); // 返回原图的副本
            }

            var croppedBitmap = new Bitmap(cropRect.Width, cropRect.Height);

            using (var graphics = Graphics.FromImage(croppedBitmap))
            {
                graphics.DrawImage(source, 0, 0, cropRect, GraphicsUnit.Pixel);
            }

            return croppedBitmap;
        }

        /// <summary>
        /// 将图片添加到PDF文档
        /// </summary>
        private void AddImageToPdf(Document document, System.Drawing.Image image)
        {
            // 将System.Drawing.Image转换为iTextSharp.text.Image
            using (var memoryStream = new MemoryStream())
            {
                image.Save(memoryStream, ImageFormat.Png);
                memoryStream.Position = 0;

                var iTextImage = iTextSharp.text.Image.GetInstance(memoryStream.ToArray());

                // 计算页面尺寸
                float pageWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                float pageHeight = document.PageSize.Height - document.TopMargin - document.BottomMargin;

                // 缩放图片以适应页面
                if (settings.PreserveAspectRatio)
                {
                    iTextImage.ScaleToFit(pageWidth, pageHeight);
                }
                else
                {
                    iTextImage.ScaleAbsolute(pageWidth, pageHeight);
                }

                // 居中显示
                iTextImage.Alignment = iTextSharp.text.Image.ALIGN_CENTER;

                // 添加新页面并插入图片
                document.NewPage();
                document.Add(iTextImage);
            }
        }

        /// <summary>
        /// 估算输出文件大小
        /// </summary>
        public long EstimateOutputSize(List<dynamic> slides)
        {
            // 简单估算：每页约50KB - 500KB，取决于DPI和质量
            int baseSize = 50 * 1024; // 50KB base
            int dpiMultiplier = settings.ExportDpi / 150; // DPI倍数
            int qualityMultiplier = settings.PdfQuality / 50; // 质量倍数

            long estimatedSizePerPage = baseSize * dpiMultiplier * qualityMultiplier;
            return estimatedSizePerPage * slides.Count;
        }
    }
}