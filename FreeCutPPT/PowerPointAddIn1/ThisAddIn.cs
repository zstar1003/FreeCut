using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        /// <summary>
        /// 获取选中的幻灯片
        /// </summary>
        public List<dynamic> GetSelectedSlides()
        {
            var selectedSlides = new List<dynamic>();

            try
            {
                PowerPoint.Selection selection = Application.ActiveWindow.Selection;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    PowerPoint.SlideRange slideRange = selection.SlideRange;
                    for (int i = 1; i <= slideRange.Count; i++)
                    {
                        selectedSlides.Add(slideRange[i]);
                    }
                }
                else if (Application.ActiveWindow.View.Type == PowerPoint.PpViewType.ppViewSlide)
                {
                    // 如果在幻灯片视图中，添加当前幻灯片
                    PowerPoint.Slide currentSlide = (PowerPoint.Slide)Application.ActiveWindow.View.Slide;
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
        public List<dynamic> GetAllSlides()
        {
            var allSlides = new List<dynamic>();

            try
            {
                PowerPoint.Slides slides = Application.ActivePresentation.Slides;

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
        public string ExportSlideToImage(dynamic slide, string outputPath, int dpi = 600)
        {
            try
            {
                PowerPoint.Slide pptSlide = (PowerPoint.Slide)slide;

                // 获取幻灯片尺寸（单位：磅 points, 1磅 = 1/72英寸）
                float slideWidthInPoints = pptSlide.Parent.PageSetup.SlideWidth;
                float slideHeightInPoints = pptSlide.Parent.PageSetup.SlideHeight;

                // 转换为英寸
                float slideWidthInInches = slideWidthInPoints / 72f;
                float slideHeightInInches = slideHeightInPoints / 72f;

                // 根据DPI计算像素尺寸
                int exportWidth = (int)Math.Round(slideWidthInInches * dpi);
                int exportHeight = (int)Math.Round(slideHeightInInches * dpi);

                // 导出为图片
                pptSlide.Export(outputPath, "PNG", exportWidth, exportHeight);
                return outputPath;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"导出幻灯片图片失败: {ex.Message}");
                throw;
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
