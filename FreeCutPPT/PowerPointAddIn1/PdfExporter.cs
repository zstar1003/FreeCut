using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;

namespace PowerPointAddIn1
{
    /// <summary>
    /// PDF导出器 - 使用Windows内置PDF功能导出真正的PDF文档
    /// </summary>
    public class PdfExporter
    {
        private readonly CropSettings settings;
        private List<System.Drawing.Image> imagesToPrint;
        private int currentPageIndex;
        private string outputFilePath;

        public PdfExporter(CropSettings settings)
        {
            this.settings = settings ?? throw new ArgumentNullException(nameof(settings));
            this.imagesToPrint = new List<System.Drawing.Image>();
        }

        /// <summary>
        /// 导出幻灯片为PDF文档
        /// </summary>
        public void ExportToPdf(List<dynamic> slides, string outputPath, Action<int, string> progressCallback = null)
        {
            if (slides == null || slides.Count == 0)
                throw new ArgumentException("幻灯片列表不能为空", nameof(slides));

            System.Diagnostics.Debug.WriteLine($"========== 开始导出 {slides.Count} 张幻灯片 ==========");
            System.Diagnostics.Debug.WriteLine($"导出DPI: {settings.ExportDpi}");
            System.Diagnostics.Debug.WriteLine($"自动检测边界: {settings.AutoDetectBounds}");
            System.Diagnostics.Debug.WriteLine($"边距: 上{settings.TopMargin} 下{settings.BottomMargin} 左{settings.LeftMargin} 右{settings.RightMargin}");

            progressCallback?.Invoke(0, "初始化PDF导出...");

            try
            {
                // 清理之前的图片
                ClearImages();

                // 处理所有幻灯片为图片
                for (int i = 0; i < slides.Count; i++)
                {
                    var slide = slides[i];
                    var progress = (int)((float)(i + 1) / slides.Count * 80); // 80% for processing
                    progressCallback?.Invoke(progress, $"正在处理第 {i + 1}/{slides.Count} 张幻灯片...");

                    try
                    {
                        var processedImage = ProcessSlideToImage(slide, i + 1);
                        if (processedImage != null)
                        {
                            imagesToPrint.Add(processedImage);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"处理第 {i + 1} 张幻灯片失败: {ex.Message}");
                        // 继续处理其他幻灯片
                    }
                }

                if (imagesToPrint.Count > 0)
                {
                    progressCallback?.Invoke(85, "正在生成PDF文档...");

                    // 使用Windows打印功能导出为PDF
                    PrintToPdf(outputPath);

                    progressCallback?.Invoke(100, "PDF导出完成");
                }
                else
                {
                    throw new InvalidOperationException("没有成功处理任何幻灯片");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"PDF导出失败: {ex.Message}", ex);
            }
            finally
            {
                // 清理图片资源
                ClearImages();
            }
        }

        /// <summary>
        /// 处理单张幻灯片为图片并返回处理后的图片
        /// </summary>
        private System.Drawing.Image ProcessSlideToImage(dynamic slide, int slideNumber)
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

                // 使用PowerPoint导出幻灯片为图片（指定像素尺寸）
                slide.Export(tempImagePath, "PNG", exportWidth, exportHeight);

                // 处理图片
                using (var originalImage = System.Drawing.Image.FromFile(tempImagePath))
                {
                    System.Diagnostics.Debug.WriteLine($"  导出后实际尺寸: {originalImage.Width}x{originalImage.Height}");

                    var croppedImage = CropImage(originalImage);

                    System.Diagnostics.Debug.WriteLine($"  最终尺寸: {croppedImage.Width}x{croppedImage.Height}");
                    return croppedImage;
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
        /// 使用Windows打印功能将图片列表导出为PDF
        /// </summary>
        private void PrintToPdf(string outputPath)
        {
            this.outputFilePath = outputPath;
            this.currentPageIndex = 0;

            System.Diagnostics.Debug.WriteLine($"========== 开始打印PDF: 共{imagesToPrint.Count}页 ==========");

            using (var printDocument = new PrintDocument())
            {
                // 设置打印机为Microsoft Print to PDF
                printDocument.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                printDocument.PrinterSettings.PrintToFile = true;
                printDocument.PrinterSettings.PrintFileName = outputPath;

                // 为第一页设置默认纸张大小
                if (imagesToPrint.Count > 0)
                {
                    var firstImage = imagesToPrint[0];
                    double dpi = settings.ExportDpi;

                    // 计算纸张尺寸（1/100英寸）
                    int paperWidth = (int)Math.Round(firstImage.Width / dpi * 100);
                    int paperHeight = (int)Math.Round(firstImage.Height / dpi * 100);

                    // 确保横向页面：PaperSize的第一个参数（宽度）应该是较大的值
                    // 如果图片是横向的（宽>高），直接使用
                    // PaperSize会自动处理方向
                    var paperSize = new PaperSize("PPTSlide", paperWidth, paperHeight);
                    printDocument.DefaultPageSettings.PaperSize = paperSize;
                    printDocument.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);

                    // 明确设置为横向或纵向
                    printDocument.DefaultPageSettings.Landscape = (firstImage.Width > firstImage.Height);

                    System.Diagnostics.Debug.WriteLine($"默认页面设置（第一页）: 纸张={paperWidth/100.0}\"x{paperHeight/100.0}\" ({firstImage.Width}px x {firstImage.Height}px @ {dpi} DPI)");
                    System.Diagnostics.Debug.WriteLine($"  Landscape={printDocument.DefaultPageSettings.Landscape}, 图片比例={firstImage.Width}:{firstImage.Height}");
                }

                // 绑定事件 - QueryPageSettings在PrintPage之前调用（从第二页开始）
                printDocument.QueryPageSettings += PrintDocument_QueryPageSettings;
                printDocument.PrintPage += PrintDocument_PrintPage;

                try
                {
                    printDocument.Print();
                    System.Diagnostics.Debug.WriteLine($"========== PDF打印完成 ==========");
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"PDF打印失败: {ex.Message}. 请确保已安装Microsoft Print to PDF驱动程序。", ex);
                }
            }
        }

        /// <summary>
        /// 查询页面设置事件 - 在每页打印前调用（从第二页开始）
        /// </summary>
        private void PrintDocument_QueryPageSettings(object sender, QueryPageSettingsEventArgs e)
        {
            // currentPageIndex在PrintPage中递增，所以这里看到的是下一页的索引
            int nextPageIndex = currentPageIndex;

            if (nextPageIndex < imagesToPrint.Count)
            {
                var image = imagesToPrint[nextPageIndex];

                // 计算页面尺寸（单位：1/100英寸）
                double dpi = settings.ExportDpi;
                int paperWidth = (int)Math.Round(image.Width / dpi * 100);
                int paperHeight = (int)Math.Round(image.Height / dpi * 100);

                // 设置页面大小
                e.PageSettings.PaperSize = new PaperSize("PPTSlide", paperWidth, paperHeight);
                e.PageSettings.Margins = new Margins(0, 0, 0, 0);

                // 设置横向或纵向
                e.PageSettings.Landscape = (image.Width > image.Height);

                System.Diagnostics.Debug.WriteLine($"QueryPageSettings 页面 {nextPageIndex + 1}: 纸张={paperWidth/100.0}\"x{paperHeight/100.0}\" ({image.Width}px x {image.Height}px @ {dpi} DPI)");
                System.Diagnostics.Debug.WriteLine($"  Landscape={e.PageSettings.Landscape}");
            }
        }

        /// <summary>
        /// 打印页面事件处理
        /// </summary>
        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            if (currentPageIndex < imagesToPrint.Count)
            {
                var image = imagesToPrint[currentPageIndex];
                var graphics = e.Graphics;

                // 设置高质量渲染模式
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                // 获取页面尺寸（单位是1/100英寸）
                var pageBoundsWidth = e.PageBounds.Width;
                var pageBoundsHeight = e.PageBounds.Height;

                // Graphics的DPI
                float graphicsDpiX = graphics.DpiX;
                float graphicsDpiY = graphics.DpiY;

                // 计算绘制区域（Graphics坐标单位）
                // PageBounds是1/100英寸，需要转换为Graphics坐标
                float drawWidth = pageBoundsWidth * graphicsDpiX / 100f;
                float drawHeight = pageBoundsHeight * graphicsDpiY / 100f;

                System.Diagnostics.Debug.WriteLine($"PrintPage 页面 {currentPageIndex + 1}:");
                System.Diagnostics.Debug.WriteLine($"  PageBounds={pageBoundsWidth}x{pageBoundsHeight} (1/100英寸)");
                System.Diagnostics.Debug.WriteLine($"  GraphicsDPI={graphicsDpiX}x{graphicsDpiY}");
                System.Diagnostics.Debug.WriteLine($"  绘制尺寸={drawWidth}x{drawHeight} (Graphics单位)");
                System.Diagnostics.Debug.WriteLine($"  图片尺寸={image.Width}x{image.Height} (像素)");
                System.Diagnostics.Debug.WriteLine($"  Landscape={e.PageSettings.Landscape}");

                // 使用计算出的绘制尺寸，填满整个页面
                graphics.DrawImage(image, 0, 0, drawWidth, drawHeight);

                currentPageIndex++;

                // 检查是否还有更多页面
                e.HasMorePages = currentPageIndex < imagesToPrint.Count;
            }
            else
            {
                e.HasMorePages = false;
            }
        }

        /// <summary>
        /// 清理图片资源
        /// </summary>
        private void ClearImages()
        {
            if (imagesToPrint != null)
            {
                foreach (var image in imagesToPrint)
                {
                    image?.Dispose();
                }
                imagesToPrint.Clear();
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

                // 应用边距
                var cropRect = new Rectangle(
                    Math.Max(0, contentBounds.Left - settings.LeftMargin),
                    Math.Max(0, contentBounds.Top - settings.TopMargin),
                    Math.Min(bitmap.Width - Math.Max(0, contentBounds.Left - settings.LeftMargin),
                            contentBounds.Width + settings.LeftMargin + settings.RightMargin),
                    Math.Min(bitmap.Height - Math.Max(0, contentBounds.Top - settings.TopMargin),
                            contentBounds.Height + settings.TopMargin + settings.BottomMargin)
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
            var cropRect = new Rectangle(
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