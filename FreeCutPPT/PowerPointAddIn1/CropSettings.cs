using System;
using System.IO;
using System.Xml.Serialization;
// using Newtonsoft.Json;

namespace PowerPointAddIn1
{
    /// <summary>
    /// 裁剪和导出设置
    /// </summary>
    public class CropSettings
    {
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "FreeCut",
            "settings.json");

        /// <summary>
        /// 是否启用自动检测边界
        /// </summary>
        public bool AutoDetectBounds { get; set; } = true;

        /// <summary>
        /// 上边距（像素）
        /// </summary>
        public int TopMargin { get; set; } = 10;

        /// <summary>
        /// 下边距（像素）
        /// </summary>
        public int BottomMargin { get; set; } = 10;

        /// <summary>
        /// 左边距（像素）
        /// </summary>
        public int LeftMargin { get; set; } = 10;

        /// <summary>
        /// 右边距（像素）
        /// </summary>
        public int RightMargin { get; set; } = 10;

        /// <summary>
        /// 检测容差（0-50）
        /// </summary>
        public int DetectionTolerance { get; set; } = 5;

        /// <summary>
        /// PDF质量（1-100）
        /// </summary>
        public int PdfQuality { get; set; } = 100;

        /// <summary>
        /// 导出DPI
        /// </summary>
        public int ExportDpi { get; set; } = 150;

        /// <summary>
        /// 是否保持宽高比
        /// </summary>
        public bool PreserveAspectRatio { get; set; } = true;

        /// <summary>
        /// 背景色检测模式
        /// </summary>
        public BackgroundDetectionMode BackgroundMode { get; set; } = BackgroundDetectionMode.AutoDetect;

        /// <summary>
        /// 自定义背景色（ARGB）
        /// </summary>
        public uint CustomBackgroundColor { get; set; } = 0xFFFFFFFF; // 白色

        /// <summary>
        /// 保存设置
        /// </summary>
        public void Save()
        {
            try
            {
                var directory = Path.GetDirectoryName(SettingsPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var serializer = new XmlSerializer(typeof(CropSettings));
                using (var writer = new StreamWriter(SettingsPath))
                {
                    serializer.Serialize(writer, this);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"保存设置失败: {ex.Message}");
                throw new InvalidOperationException($"保存设置失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 加载设置
        /// </summary>
        public static CropSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var serializer = new XmlSerializer(typeof(CropSettings));
                    using (var reader = new StreamReader(SettingsPath))
                    {
                        var settings = (CropSettings)serializer.Deserialize(reader);
                        return settings ?? new CropSettings();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"加载设置失败: {ex.Message}");
            }

            return new CropSettings();
        }

        /// <summary>
        /// 重置为默认值
        /// </summary>
        public void ResetToDefault()
        {
            AutoDetectBounds = true;
            TopMargin = 10;
            BottomMargin = 10;
            LeftMargin = 10;
            RightMargin = 10;
            DetectionTolerance = 5;
            PdfQuality = 100;
            ExportDpi = 150;
            PreserveAspectRatio = true;
            BackgroundMode = BackgroundDetectionMode.AutoDetect;
            CustomBackgroundColor = 0xFFFFFFFF;
        }

        /// <summary>
        /// 验证设置值
        /// </summary>
        public bool Validate(out string errorMessage)
        {
            errorMessage = string.Empty;

            // 边距范围检查
            if (TopMargin < 0 || TopMargin > 500)
            {
                errorMessage = "上边距必须在0-500像素之间";
                return false;
            }

            if (BottomMargin < 0 || BottomMargin > 500)
            {
                errorMessage = "下边距必须在0-500像素之间";
                return false;
            }

            if (LeftMargin < 0 || LeftMargin > 500)
            {
                errorMessage = "左边距必须在0-500像素之间";
                return false;
            }

            if (RightMargin < 0 || RightMargin > 500)
            {
                errorMessage = "右边距必须在0-500像素之间";
                return false;
            }

            // 检测容差检查
            if (DetectionTolerance < 0 || DetectionTolerance > 50)
            {
                errorMessage = "检测容差必须在0-50之间";
                return false;
            }

            // PDF质量检查
            if (PdfQuality < 1 || PdfQuality > 100)
            {
                errorMessage = "PDF质量必须在1-100之间";
                return false;
            }

            // DPI检查
            var validDpis = new[] { 72, 150, 300, 600 };
            bool validDpi = false;
            foreach (var dpi in validDpis)
            {
                if (ExportDpi == dpi)
                {
                    validDpi = true;
                    break;
                }
            }

            if (!validDpi)
            {
                errorMessage = "DPI必须选择72、150、300或600";
                return false;
            }

            return true;
        }

        /// <summary>
        /// 简单验证（向后兼容）
        /// </summary>
        public bool Validate()
        {
            string errorMessage;
            return Validate(out errorMessage);
        }

        /// <summary>
        /// 复制设置
        /// </summary>
        public CropSettings Clone()
        {
            return new CropSettings
            {
                AutoDetectBounds = this.AutoDetectBounds,
                TopMargin = this.TopMargin,
                BottomMargin = this.BottomMargin,
                LeftMargin = this.LeftMargin,
                RightMargin = this.RightMargin,
                DetectionTolerance = this.DetectionTolerance,
                PdfQuality = this.PdfQuality,
                ExportDpi = this.ExportDpi,
                PreserveAspectRatio = this.PreserveAspectRatio,
                BackgroundMode = this.BackgroundMode,
                CustomBackgroundColor = this.CustomBackgroundColor
            };
        }

        /// <summary>
        /// 获取DPI选项
        /// </summary>
        public static int[] GetAvailableDpis()
        {
            return new[] { 72, 150, 300, 600 };
        }

        /// <summary>
        /// 获取DPI描述
        /// </summary>
        public static string GetDpiDescription(int dpi)
        {
            switch (dpi)
            {
                case 72:
                    return "72 DPI (网页显示)";
                case 150:
                    return "150 DPI (一般打印)";
                case 300:
                    return "300 DPI (高质量印刷)";
                case 600:
                    return "600 DPI (专业印刷)";
                default:
                    return $"{dpi} DPI";
            }
        }
    }

    /// <summary>
    /// 背景检测模式
    /// </summary>
    public enum BackgroundDetectionMode
    {
        /// <summary>
        /// 自动检测（使用幻灯片角落颜色）
        /// </summary>
        AutoDetect,

        /// <summary>
        /// 使用白色背景
        /// </summary>
        White,

        /// <summary>
        /// 使用透明背景
        /// </summary>
        Transparent,

        /// <summary>
        /// 自定义颜色
        /// </summary>
        Custom
    }
}