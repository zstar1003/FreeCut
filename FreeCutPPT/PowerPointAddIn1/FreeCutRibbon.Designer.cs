
namespace PowerPointAddIn1
{
    partial class FreeCutRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public FreeCutRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupMain = this.Factory.CreateRibbonGroup();
            this.btnSettings = this.Factory.CreateRibbonButton();
            this.btnExportPDF = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupMain.SuspendLayout();
            this.SuspendLayout();
            //
            // tab1
            //
            this.tab1.Groups.Add(this.groupMain);
            this.tab1.Label = "FreeCut";
            this.tab1.Name = "tab1";
            //
            // groupMain
            //
            this.groupMain.Items.Add(this.btnSettings);
            this.groupMain.Items.Add(this.btnExportPDF);
            this.groupMain.Label = "PDF导出";
            this.groupMain.Name = "groupMain";
            //
            // btnSettings
            //
            this.btnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSettings.Label = "FreeCut设置";
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.OfficeImageId = "PageSetupDialog";
            this.btnSettings.ShowImage = true;
            this.btnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSettings_Click);
            //
            // btnExportPDF
            //
            this.btnExportPDF.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportPDF.Label = "导出PDF";
            this.btnExportPDF.Name = "btnExportPDF";
            this.btnExportPDF.OfficeImageId = "FileSaveAsPdf";
            this.btnExportPDF.ShowImage = true;
            this.btnExportPDF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportPDF_Click);
            //
            // FreeCutRibbon
            //
            this.Name = "FreeCutRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.FreeCutRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupMain.ResumeLayout(false);
            this.groupMain.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMain;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportPDF;
    }

    partial class ThisRibbonCollection
    {
        internal FreeCutRibbon FreeCutRibbon
        {
            get { return this.GetRibbon<FreeCutRibbon>(); }
        }
    }
}
