using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FreeCut
{
    [ComVisible(true)]
    [Guid("87654321-4321-4321-4321-CBA987654321")]
    [ProgId("FreeCut.FreeCutRibbon")]
    public class FreeCutRibbon
    {
        private SettingsForm settingsForm;
        private ProgressForm progressForm;
        private PreviewForm previewForm;

        public FreeCutRibbon()
        {

        }

        /// <summary>
        /// 获取Ribbon XML
        /// </summary>
        public string GetCustomUI(string ribbonID)
        {
            return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='FreeCutTab' label='FreeCut'>
        <group id='FreeCutGroup' label='PDF导出'>
          <button id='FreeCutSettings'
                  label='FreeCut设置'
                  size='large'
                  imageMso='FileExportMenu'
                  onAction='OnFreeCutSettings'
                  screentip='打开FreeCut设置面板'
                  supertip='配置裁剪边距、PDF质量等导出参数' />
          <button id='ExportPDF'
                  label='导出PDF'
                  size='large'
                  imageMso='ExportPDF'
                  onAction='OnExportPDF'
                  screentip='快速导出PDF'
                  supertip='使用当前设置导出选中的幻灯片为PDF文件' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        /// <summary>
        /// 打开设置窗口
        /// </summary>
        public void OnFreeCutSettings(object control)
        {
            try
            {
                if (settingsForm == null || settingsForm.IsDisposed)
                {
                    settingsForm = new SettingsForm();
                }

                if (!settingsForm.Visible)
                {
                    settingsForm.Show();
                }
                else
                {
                    settingsForm.BringToFront();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开设置窗口失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 导出PDF
        /// </summary>
        public void OnExportPDF(object control)
        {
            try
            {
                var selectedSlides = ThisAddIn.GetSelectedSlides();
                if (selectedSlides.Count == 0)
                {
                    MessageBox.Show("请先选择要导出的幻灯片", "提示",
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
                        StartExportProcess(selectedSlides, saveDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出PDF失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 开始导出流程
        /// </summary>
        private void StartExportProcess(List<dynamic> slides, string outputPath)
        {
            try
            {
                // 显示进度窗口
                if (progressForm == null || progressForm.IsDisposed)
                {
                    progressForm = new ProgressForm();
                }

                progressForm.Show();
                progressForm.SetProgressText("准备导出...");
                progressForm.SetProgress(0);

                // 加载设置
                var settings = CropSettings.Load();

                // 创建PDF导出器
                var exporter = new PdfExporter(settings);

                // 执行导出
                exporter.ExportToPdf(slides, outputPath, (progress, message) =>
                {
                    progressForm.SetProgress(progress);
                    progressForm.SetProgressText(message);
                });

                progressForm.SetProgress(100);
                progressForm.SetProgressText("导出完成！");

                MessageBox.Show($"PDF导出成功！\n保存位置：{outputPath}", "成功",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                progressForm.Hide();
            }
            catch (Exception ex)
            {
                progressForm?.Hide();
                throw;
            }
        }

        /// <summary>
        /// 显示预览窗口
        /// </summary>
        public void ShowPreview()
        {
            try
            {
                var selectedSlides = ThisAddIn.GetSelectedSlides();
                if (selectedSlides.Count == 0)
                {
                    MessageBox.Show("请先选择要预览的幻灯片", "提示",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (previewForm == null || previewForm.IsDisposed)
                {
                    previewForm = new PreviewForm();
                }

                previewForm.LoadPreview(selectedSlides);
                previewForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"预览失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            try
            {
                settingsForm?.Close();
                settingsForm?.Dispose();
                settingsForm = null;

                progressForm?.Close();
                progressForm?.Dispose();
                progressForm = null;

                previewForm?.Close();
                previewForm?.Dispose();
                previewForm = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"FreeCutRibbon释放资源时发生错误: {ex.Message}");
            }
        }
    }
}