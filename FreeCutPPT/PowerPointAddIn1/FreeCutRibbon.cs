using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class FreeCutRibbon
    {
        private SettingsForm settingsForm;
        private ProgressForm progressForm;
        private PreviewForm previewForm;

        private void FreeCutRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnSettings_Click(object sender, RibbonControlEventArgs e)
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

        private void btnExportPDF_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var selectedSlides = Globals.ThisAddIn.GetSelectedSlides();
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

        private void btnPreview_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var selectedSlides = Globals.ThisAddIn.GetSelectedSlides();
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

        private void btnRefresh_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                MessageBox.Show("设置已重新加载", "FreeCut",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"重新加载设置失败：{ex.Message}", "错误",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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
                progressForm.ShowCompleted();
            }
            catch
            {
                progressForm?.Hide();
                throw;
            }
        }
    }
}
