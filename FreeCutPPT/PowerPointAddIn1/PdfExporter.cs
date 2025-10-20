using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using SkiaSharp;

namespace PowerPointAddIn1
{
    /// <summary>
    /// PDF导出器 - 使用SkiaSharp直接生成PDF文档，避免Windows打印系统的坐标转换问题
    /// </summary>
    public class PdfExporter
    {
        private readonly CropSettings settings;

        public PdfExporter(CropSettings settings)
        {
            this.settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <summary>
        /// 导出幻灯片为PDF文档
        /// </summary>
        public void ExportToPdf(List<dynamic> slides, string outputPath, Action<int, string> progressCallback = null)
        {
            if (slides == null || slides.Count == 0)
                throw new ArgumentException("幻灯片列表不能为空", nameof(slides));

            System.Diagnostics.Debug.WriteLine($"========== 开始导出 {slides.Count} 张幻灯片 (SkiaSharp) ==========");
            System.Diagnostics.Debug.WriteLine($"导出DPI: {settings.ExportDpi}");
            System.Diagnostics.Debug.WriteLine($"自动检测边界: {settings.AutoDetectBounds}");
            System.Diagnostics.Debug.WriteLine($"边距: 上{settings.TopMargin} 下{settings.BottomMargin} 左{settings.LeftMargin} 右{settings.RightMargin}");

            progressCallback?.Invoke(0, "初始化PDF导出...");

            try
            {
                // 先处理所有幻灯片为图片
                var processedImages = new List<(System.Drawing.Image image, float width, float height)>();

                try
                {
                    for (int i = 0; i < slides.Count; i++)
                    {
                        var slide = slides[i];
                        var progress = (int)((float)(i + 1) / slides.Count * 80);
                        progressCallback?.Invoke(progress, $"正在处理第 {i + 1}/{slides.Count} 张幻灯片...");

                        try
                        {
                            var imageData = ProcessSlideToImageData(slide, i + 1);
                            processedImages.Add(imageData);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"处理第 {i + 1} 张幻灯片失败: {ex.Message}");
                            throw;
                        }
                    }

                    progressCallback?.Invoke(85, "正在生成PDF文档...");

                    // 一次性创建包含所有页面的PDF
                    CreatePdfDocument(processedImages, outputPath);

                    progressCallback?.Invoke(100, "PDF导出完成");
                    System.Diagnostics.Debug.WriteLine($"========== PDF导出完成 ==========");
                }
                finally
                {
                    // 清理所有图片资源
                    foreach (var (image, _, _) in processedImages)
                    {
                        image?.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"PDF导出失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 处理单张幻灯片为图片数据
        /// </summary>
        private (System.Drawing.Image image, float widthInPoints, float heightInPoints) ProcessSlideToImageData(dynamic slide, int slideNumber)
        {
            string tempImagePath = Path.Combine(Path.GetTempPath(), $"FreeCut_Slide_{slideNumber}_{Guid.NewGuid()}.png");

            try
            {
                // 获取幻灯片尺寸（单位：磅 points, 1磅 = 1/72英寸）
                float slideWidthInPoints = slide.Parent.PageSetup.SlideWidth;
                float slideHeightInPoints = slide.Parent.PageSetup.SlideHeight;

                // 转换为英寸
                float slideWidthInInches = slideWidthInPoints / 72f;
                float slideHeightInInches = slideHeightInPoints / 72f;

                // 根据DPI计算像素尺寸
                int exportWidth = (int)Math.Round(slideWidthInInches * settings.ExportDpi);
                int exportHeight = (int)Math.Round(slideHeightInInches * settings.ExportDpi);

                System.Diagnostics.Debug.WriteLine($"处理幻灯片 {slideNumber}:");
                System.Diagnostics.Debug.WriteLine($"  幻灯片尺寸: {slideWidthInPoints:F1} x {slideHeightInPoints:F1} 磅 = {slideWidthInInches:F2}\" x {slideHeightInInches:F2}\"");
                System.Diagnostics.Debug.WriteLine($"  导出尺寸: {exportWidth} x {exportHeight} 像素 @ {settings.ExportDpi} DPI");

                // 使用PowerPoint导出幻灯片为图片
                slide.Export(tempImagePath, "PNG", exportWidth, exportHeight);

                // 加载并处理图片
                using (var originalImage = System.Drawing.Image.FromFile(tempImagePath))
                {
                    System.Diagnostics.Debug.WriteLine($"  导出后实际尺寸: {originalImage.Width}x{originalImage.Height}");

                    // 裁剪图片
                    var croppedImage = CropImage(originalImage);
                    System.Diagnostics.Debug.WriteLine($"  最终尺寸: {croppedImage.Width}x{croppedImage.Height}");

                    // 计算PDF页面的物理尺寸（英寸）
                    float pageWidthInInches = croppedImage.Width / (float)settings.ExportDpi;
                    float pageHeightInInches = croppedImage.Height / (float)settings.ExportDpi;

                    // 转换为点（1英寸 = 72点）
                    float pageWidthInPoints = pageWidthInInches * 72f;
                    float pageHeightInPoints = pageHeightInInches * 72f;

                    System.Diagnostics.Debug.WriteLine($"  PDF页面尺寸: {pageWidthInPoints:F1} x {pageHeightInPoints:F1} 点 ({pageWidthInInches:F2}\" x {pageHeightInInches:F2}\")");

                    return (croppedImage, pageWidthInPoints, pageHeightInPoints);
                }
            }
            finally
            {
                // 清理临时文件
                if (File.Exists(tempImagePath))
                {
                    try { File.Delete(tempImagePath); } catch { }
                }
            }
        }

        /// <summary>
        /// 使用SkiaSharp创建包含所有页面的PDF文档
        /// </summary>
        private void CreatePdfDocument(List<(System.Drawing.Image image, float width, float height)> pages, string outputPath)
        {
            using (var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            using (var document = SKDocument.CreatePdf(stream))
            {
                foreach (var (image, pageWidth, pageHeight) in pages)
                {
                    // 将 System.Drawing.Image 转换为 SKBitmap
                    using (var memoryStream = new MemoryStream())
                    {
                        image.Save(memoryStream, ImageFormat.Png);
                        memoryStream.Position = 0;

                        using (var skBitmap = SKBitmap.Decode(memoryStream))
                        {
                            // 创建页面（尺寸单位：点）
                            using (var canvas = document.BeginPage(pageWidth, pageHeight))
                            {
                                // 绘制图片，填满整个页面
                                var destRect = new SKRect(0, 0, pageWidth, pageHeight);
                                canvas.DrawBitmap(skBitmap, destRect);
                            }
                            document.EndPage();
                        }
                    }
                }

                // 关闭文档
                document.Close();
            }
        }

        /// <summary>
        /// 裁剪图片
        /// </summary>
        private System.Drawing.Image CropImage(System.Drawing.Image originalImage)
        {
            System.Diagnostics.Debug.WriteLine($"  裁剪前尺寸: {originalImage.Width}x{originalImage.Height}");

            System.Drawing.Image result;
            if (settings.AutoDetectBounds)
            {
                System.Diagnostics.Debug.WriteLine($"  使用自动检测边界");
                result = CropImageWithAutoDetection(originalImage);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"  使用固定边距裁剪");
                result = CropImageWithFixedMargins(originalImage);
            }

            System.Diagnostics.Debug.WriteLine($"  裁剪后尺寸: {result.Width}x{result.Height}");
            return result;
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

                System.Diagnostics.Debug.WriteLine($"  检测到内容边界: Left={contentBounds.Left}, Top={contentBounds.Top}, Width={contentBounds.Width}, Height={contentBounds.Height}");

                // 边距需要根据DPI缩放，以保持相同的物理尺寸
                const int baseDpi = 150;
                double dpiScale = (double)settings.ExportDpi / baseDpi;

                // 缩放边距
                int scaledLeftMargin = (int)Math.Round(settings.LeftMargin * dpiScale);
                int scaledRightMargin = (int)Math.Round(settings.RightMargin * dpiScale);
                int scaledTopMargin = (int)Math.Round(settings.TopMargin * dpiScale);
                int scaledBottomMargin = (int)Math.Round(settings.BottomMargin * dpiScale);

                System.Diagnostics.Debug.WriteLine($"  边距缩放: DPI {settings.ExportDpi} / 基准 {baseDpi} = {dpiScale:F2}x");
                System.Diagnostics.Debug.WriteLine($"  缩放后边距: L{scaledLeftMargin} R{scaledRightMargin} T{scaledTopMargin} B{scaledBottomMargin}");

                // 应用边距
                var cropRect = new Rectangle(
                    Math.Max(0, contentBounds.Left - scaledLeftMargin),
                    Math.Max(0, contentBounds.Top - scaledTopMargin),
                    Math.Min(bitmap.Width - Math.Max(0, contentBounds.Left - scaledLeftMargin),
                            contentBounds.Width + scaledLeftMargin + scaledRightMargin),
                    Math.Min(bitmap.Height - Math.Max(0, contentBounds.Top - scaledTopMargin),
                            contentBounds.Height + scaledTopMargin + scaledBottomMargin)
                );

                System.Diagnostics.Debug.WriteLine($"  应用边距后裁剪区域: {cropRect}");

                return CropBitmap(bitmap, cropRect);
            }
        }

        /// <summary>
        /// 使用固定边距裁剪
        /// </summary>
        private System.Drawing.Image CropImageWithFixedMargins(System.Drawing.Image originalImage)
        {
            // 边距需要根据DPI缩放，以保持相同的物理尺寸
            // 假设边距是基于150 DPI设置的
            const int baseDpi = 150;
            double dpiScale = (double)settings.ExportDpi / baseDpi;

            // 缩放边距
            int scaledLeftMargin = (int)Math.Round(settings.LeftMargin * dpiScale);
            int scaledRightMargin = (int)Math.Round(settings.RightMargin * dpiScale);
            int scaledTopMargin = (int)Math.Round(settings.TopMargin * dpiScale);
            int scaledBottomMargin = (int)Math.Round(settings.BottomMargin * dpiScale);

            System.Diagnostics.Debug.WriteLine($"  边距缩放: DPI {settings.ExportDpi} / 基准 {baseDpi} = {dpiScale:F2}x");
            System.Diagnostics.Debug.WriteLine($"  原始边距: L{settings.LeftMargin} R{settings.RightMargin} T{settings.TopMargin} B{settings.BottomMargin}");
            System.Diagnostics.Debug.WriteLine($"  缩放后边距: L{scaledLeftMargin} R{scaledRightMargin} T{scaledTopMargin} B{scaledBottomMargin}");

            var cropRect = new Rectangle(
                scaledLeftMargin,
                scaledTopMargin,
                originalImage.Width - scaledLeftMargin - scaledRightMargin,
                originalImage.Height - scaledTopMargin - scaledBottomMargin
            );

            using (var bitmap = new Bitmap(originalImage))
            {
                return CropBitmap(bitmap, cropRect);
            }
        }

        /// <summary>
        /// 检测内容边界
        /// </summary>
        private Rectangle DetectContentBounds(Bitmap bitmap)
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
                return new Rectangle(0, 0, bitmap.Width, bitmap.Height);
            }

            return new Rectangle(left, top, right - left + 1, bottom - top + 1);
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
        private System.Drawing.Image CropBitmap(Bitmap source, Rectangle cropRect)
        {
            // 确保裁剪区域在图片范围内
            cropRect = Rectangle.Intersect(cropRect,
                new Rectangle(0, 0, source.Width, source.Height));

            if (cropRect.IsEmpty)
            {
                return new Bitmap(source); // 返回原图的副本
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
        /// 估算PDF输出文件大小
        /// </summary>
        public long EstimateOutputSize(List<dynamic> slides)
        {
            // PDF文件大小估算：每页约100KB - 1MB，取决于DPI和图片质量
            int baseSize = 100 * 1024; // 100KB base for PDF
            int dpiMultiplier = settings.ExportDpi / 150; // DPI倍数

            long estimatedSizePerPage = baseSize * dpiMultiplier;
            return estimatedSizePerPage * slides.Count;
        }
    }
}
