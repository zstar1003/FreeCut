using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FreeCut
{
    [ComVisible(true)]
    [Guid("12345678-1234-1234-1234-123456789ABC")]
    [ProgId("FreeCut.ThisAddIn")]
    public partial class ThisAddIn
    {
        private FreeCutRibbon ribbon;
        private static object powerPointApp;

        public static object PowerPointApplication => powerPointApp;

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            try
            {
                powerPointApp = application;

                // 初始化Ribbon
                ribbon = new FreeCutRibbon();

                System.Diagnostics.Debug.WriteLine("FreeCut插件已成功加载");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"FreeCut插件加载失败: {ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnDisconnection(int removeMode, ref Array custom)
        {
            try
            {
                ribbon?.Dispose();
                ribbon = null;
                powerPointApp = null;

                System.Diagnostics.Debug.WriteLine("FreeCut插件已卸载");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"FreeCut插件卸载时发生错误: {ex.Message}");
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
            // 插件更新时调用
        }

        public void OnStartupComplete(ref Array custom)
        {
            // PowerPoint启动完成后调用
        }

        public void OnBeginShutdown(ref Array custom)
        {
            // PowerPoint关闭前调用
        }

        /// <summary>
        /// 获取选中的幻灯片
        /// </summary>
        public static List<dynamic> GetSelectedSlides()
        {
            var selectedSlides = new List<dynamic>();

            try
            {
                if (powerPointApp == null) return selectedSlides;

                dynamic app = powerPointApp;
                dynamic activeWindow = app.ActiveWindow;

                if (activeWindow?.Selection?.Type == 2) // ppSelectionSlides = 2
                {
                    dynamic slideRange = activeWindow.Selection.SlideRange;
                    for (int i = 1; i <= slideRange.Count; i++)
                    {
                        selectedSlides.Add(slideRange[i]);
                    }
                }
                else if (activeWindow?.View?.Type == 9) // ppViewSlide = 9
                {
                    // 如果在幻灯片视图中，添加当前幻灯片
                    dynamic currentSlide = activeWindow.View.Slide;
                    selectedSlides.Add(currentSlide);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取选中幻灯片失败: {ex.Message}");
            }

            return selectedSlides;
        }

        /// <summary>
        /// 获取演示文稿中的所有幻灯片
        /// </summary>
        public static List<dynamic> GetAllSlides()
        {
            var allSlides = new List<dynamic>();

            try
            {
                if (powerPointApp == null) return allSlides;

                dynamic app = powerPointApp;
                dynamic presentation = app.ActivePresentation;
                dynamic slides = presentation.Slides;

                for (int i = 1; i <= slides.Count; i++)
                {
                    allSlides.Add(slides[i]);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取所有幻灯片失败: {ex.Message}");
            }

            return allSlides;
        }

        /// <summary>
        /// 导出幻灯片为图片
        /// </summary>
        public static string ExportSlideToImage(dynamic slide, string outputPath, int dpi = 300)
        {
            try
            {
                slide.Export(outputPath, "PNG", dpi, dpi);
                return outputPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"导出幻灯片图片失败: {ex.Message}");
                throw;
            }
        }
    }
}