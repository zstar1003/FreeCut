using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;

namespace FreeCut
{
    public partial class MainWindow : Window
    {
        private PowerPointService powerPointService;
        private CropSettings settings;
        private ObservableCollection<SlideInfo> slides;
        private PdfExporter pdfExporter;

        public MainWindow()
        {
            InitializeComponent();
            InitializeApplication();
        }

        private void InitializeApplication()
        {
            powerPointService = new PowerPointService();
            settings = CropSettings.Load();
            slides = new ObservableCollection<SlideInfo>();
            pdfExporter = new PdfExporter();

            listSlides.ItemsSource = slides;
            LoadSettings();
            UpdateUI();

            // 绑定滑块事件
            sliderQuality.ValueChanged += (s, e) => lblQuality.Text = ((int)e.NewValue).ToString();
        }

        private void LoadSettings()
        {
            txtTopMargin.Text = settings.TopMargin.ToString();
            txtBottomMargin.Text = settings.BottomMargin.ToString();
            txtLeftMargin.Text = settings.LeftMargin.ToString();
            txtRightMargin.Text = settings.RightMargin.ToString();
            txtTolerance.Text = settings.DetectionTolerance.ToString();
            chkAutoDetect.IsChecked = settings.AutoDetectBounds;
            sliderQuality.Value = settings.PdfQuality;
            cmbDPI.SelectedItem = cmbDPI.Items.Cast<ComboBoxItem>()
                .FirstOrDefault(item => item.Content.ToString() == settings.ExportDpi.ToString());
            chkPreserveRatio.IsChecked = settings.PreserveAspectRatio;
        }

        private void SaveSettings()
        {
            if (int.TryParse(txtTopMargin.Text, out int top)) settings.TopMargin = top;
            if (int.TryParse(txtBottomMargin.Text, out int bottom)) settings.BottomMargin = bottom;
            if (int.TryParse(txtLeftMargin.Text, out int left)) settings.LeftMargin = left;
            if (int.TryParse(txtRightMargin.Text, out int right)) settings.RightMargin = right;
            if (int.TryParse(txtTolerance.Text, out int tolerance)) settings.DetectionTolerance = tolerance;

            settings.AutoDetectBounds = chkAutoDetect.IsChecked ?? true;
            settings.PdfQuality = (int)sliderQuality.Value;
            settings.PreserveAspectRatio = chkPreserveRatio.IsChecked ?? true;

            if (cmbDPI.SelectedItem is ComboBoxItem selectedDPI)
            {
                if (int.TryParse(selectedDPI.Content.ToString(), out int dpi))
                {
                    settings.ExportDpi = dpi;
                }
            }
        }

        private void UpdateUI()
        {
            bool hasFile = powerPointService.IsFileLoaded;
            bool hasSlides = slides.Count > 0;
            bool hasSelection = listSlides.SelectedItems.Count > 0;

            btnRefresh.IsEnabled = hasFile;
            btnSelectAll.IsEnabled = hasSlides;
            btnSelectNone.IsEnabled = hasSlides;
            btnPreview.IsEnabled = hasSelection;
            btnExportPDF.IsEnabled = hasSelection;
            cmbPreviewSlide.IsEnabled = hasSelection;

            lblSlideCount.Text = $"共 {slides.Count} 张";
            lblSelectedCount.Text = $"已选择 {listSlides.SelectedItems.Count} 张幻灯片";

            // 更新预览下拉框
            cmbPreviewSlide.Items.Clear();
            foreach (var item in listSlides.SelectedItems.Cast<SlideInfo>())
            {
                cmbPreviewSlide.Items.Add(item);
            }
            if (cmbPreviewSlide.Items.Count > 0)
            {
                cmbPreviewSlide.SelectedIndex = 0;
            }
        }

        private async void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "选择PowerPoint文件",
                Filter = "PowerPoint文件 (*.pptx;*.ppt)|*.pptx;*.ppt|所有文件 (*.*)|*.*",
                CheckFileExists = true
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    lblStatus.Text = "正在加载PowerPoint文件...";

                    await Task.Run(() =>
                    {
                        powerPointService.LoadPresentation(dialog.FileName);
                    });

                    Dispatcher.Invoke(() =>
                    {
                        lblFileName.Text = Path.GetFileName(dialog.FileName);
                        RefreshSlideList();
                        lblStatus.Text = $"成功加载文件: {Path.GetFileName(dialog.FileName)}";
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"加载文件失败: {ex.Message}", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    lblStatus.Text = "加载文件失败";
                }
            }
        }

        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshSlideList();
        }

        private void RefreshSlideList()
        {
            try
            {
                slides.Clear();
                var slideInfos = powerPointService.GetSlideInfos();

                foreach (var slideInfo in slideInfos)
                {
                    slides.Add(slideInfo);
                }

                UpdateUI();
                lblStatus.Text = $"刷新完成，共 {slides.Count} 张幻灯片";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"刷新幻灯片列表失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnSelectAll_Click(object sender, RoutedEventArgs e)
        {
            listSlides.SelectAll();
        }

        private void BtnSelectNone_Click(object sender, RoutedEventArgs e)
        {
            listSlides.UnselectAll();
        }

        private void ListSlides_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateUI();
        }

        private void CmbPreviewSlide_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbPreviewSlide.SelectedItem is SlideInfo slideInfo)
            {
                _ = UpdatePreviewAsync(slideInfo);
            }
        }

        private async Task UpdatePreviewAsync(SlideInfo slideInfo)
        {
            try
            {
                lblStatus.Text = "正在生成预览...";

                SaveSettings(); // 确保使用最新设置

                var previewImage = await Task.Run(() =>
                    powerPointService.GetSlidePreview(slideInfo.Index, settings));

                if (previewImage != null)
                {
                    imgPreview.Source = previewImage;
                    lblStatus.Text = "预览已更新";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成预览失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                lblStatus.Text = "预览失败";
            }
        }

        private async void BtnPreview_Click(object sender, RoutedEventArgs e)
        {
            if (cmbPreviewSlide.SelectedItem is SlideInfo slideInfo)
            {
                await UpdatePreviewAsync(slideInfo);
            }
        }

        private async void BtnExportPDF_Click(object sender, RoutedEventArgs e)
        {
            if (listSlides.SelectedItems.Count == 0)
            {
                MessageBox.Show("请先选择要导出的幻灯片。", "提示",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "保存PDF文件",
                Filter = "PDF文件 (*.pdf)|*.pdf",
                DefaultExt = "pdf",
                FileName = $"FreeCut导出_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    lblStatus.Text = "正在导出PDF...";
                    SaveSettings();

                    var selectedSlides = listSlides.SelectedItems.Cast<SlideInfo>().ToList();

                    var progressWindow = new ProgressWindow();
                    progressWindow.Owner = this;
                    progressWindow.Show();

                    await Task.Run(() =>
                    {
                        pdfExporter.ExportSlides(
                            powerPointService,
                            selectedSlides,
                            dialog.FileName,
                            settings,
                            progressWindow);
                    });

                    progressWindow.Close();

                    MessageBox.Show($"PDF导出成功！\n保存位置: {dialog.FileName}", "导出完成",
                        MessageBoxButton.OK, MessageBoxImage.Information);

                    lblStatus.Text = "PDF导出完成";
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出PDF失败: {ex.Message}", "错误",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    lblStatus.Text = "PDF导出失败";
                }
            }
        }

        private void BtnSaveSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveSettings();
                settings.Save();
                MessageBox.Show("设置已保存。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                lblStatus.Text = "设置已保存";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败: {ex.Message}", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnResetSettings_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("确定要重置为默认设置吗？", "确认重置",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                settings.ResetToDefault();
                LoadSettings();
                lblStatus.Text = "设置已重置";
            }
        }

        private void ChkAutoDetect_CheckedChanged(object sender, RoutedEventArgs e)
        {
            bool isEnabled = !(chkAutoDetect.IsChecked ?? false);
            txtTopMargin.IsEnabled = isEnabled;
            txtBottomMargin.IsEnabled = isEnabled;
            txtLeftMargin.IsEnabled = isEnabled;
            txtRightMargin.IsEnabled = isEnabled;
        }

        private void BtnAbout_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "FreeCut - PPT自动裁剪PDF导出工具\n\n" +
                "版本: 1.0.0\n" +
                "功能: 智能裁剪PowerPoint幻灯片并导出为PDF\n\n" +
                "特性:\n" +
                "• 自动检测内容边界\n" +
                "• 自定义边距设置\n" +
                "• 高质量PDF导出\n" +
                "• 批量处理支持\n" +
                "• 实时预览效果\n\n" +
                "使用方法:\n" +
                "1. 点击'打开PPT文件'选择PowerPoint文件\n" +
                "2. 在幻灯片列表中选择要导出的幻灯片\n" +
                "3. 调整裁剪和导出设置\n" +
                "4. 点击'预览'查看效果\n" +
                "5. 点击'导出PDF'保存文件",
                "关于 FreeCut",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void BtnHelp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (File.Exists("README.md"))
                {
                    System.Diagnostics.Process.Start("README.md");
                }
                else
                {
                    MessageBox.Show(
                        "帮助文档未找到。\n\n" +
                        "请访问项目网站查看详细使用说明。",
                        "帮助",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch
            {
                MessageBox.Show("无法打开帮助文档。", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            try
            {
                SaveSettings();
                settings.Save();
                powerPointService?.Dispose();
            }
            catch
            {
                // 忽略关闭时的错误
            }

            base.OnClosed(e);
        }
    }
}