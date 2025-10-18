namespace FreeCut
{
    partial class SettingsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.grpMargins = new System.Windows.Forms.GroupBox();
            this.btnSetAllMargins = new System.Windows.Forms.Button();
            this.numRightMargin = new System.Windows.Forms.NumericUpDown();
            this.numLeftMargin = new System.Windows.Forms.NumericUpDown();
            this.numBottomMargin = new System.Windows.Forms.NumericUpDown();
            this.numTopMargin = new System.Windows.Forms.NumericUpDown();
            this.lblRightMargin = new System.Windows.Forms.Label();
            this.lblLeftMargin = new System.Windows.Forms.Label();
            this.lblBottomMargin = new System.Windows.Forms.Label();
            this.lblTopMargin = new System.Windows.Forms.Label();
            this.grpDetection = new System.Windows.Forms.GroupBox();
            this.lblMarginHint = new System.Windows.Forms.Label();
            this.panelColorPreview = new System.Windows.Forms.Panel();
            this.btnCustomColor = new System.Windows.Forms.Button();
            this.cmbBackgroundMode = new System.Windows.Forms.ComboBox();
            this.lblBackgroundMode = new System.Windows.Forms.Label();
            this.numTolerance = new System.Windows.Forms.NumericUpDown();
            this.lblTolerance = new System.Windows.Forms.Label();
            this.chkAutoDetect = new System.Windows.Forms.CheckBox();
            this.grpExport = new System.Windows.Forms.GroupBox();
            this.cmbDpi = new System.Windows.Forms.ComboBox();
            this.lblDpi = new System.Windows.Forms.Label();
            this.chkPreserveAspectRatio = new System.Windows.Forms.CheckBox();
            this.numPdfQuality = new System.Windows.Forms.NumericUpDown();
            this.lblPdfQuality = new System.Windows.Forms.Label();
            this.grpActions = new System.Windows.Forms.GroupBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnPreview = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.lblValidation = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.grpMargins.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numRightMargin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLeftMargin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numBottomMargin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTopMargin)).BeginInit();
            this.grpDetection.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numTolerance)).BeginInit();
            this.grpExport.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numPdfQuality)).BeginInit();
            this.grpActions.SuspendLayout();
            this.SuspendLayout();
            //
            // grpMargins
            //
            this.grpMargins.Controls.Add(this.btnSetAllMargins);
            this.grpMargins.Controls.Add(this.numRightMargin);
            this.grpMargins.Controls.Add(this.numLeftMargin);
            this.grpMargins.Controls.Add(this.numBottomMargin);
            this.grpMargins.Controls.Add(this.numTopMargin);
            this.grpMargins.Controls.Add(this.lblRightMargin);
            this.grpMargins.Controls.Add(this.lblLeftMargin);
            this.grpMargins.Controls.Add(this.lblBottomMargin);
            this.grpMargins.Controls.Add(this.lblTopMargin);
            this.grpMargins.Location = new System.Drawing.Point(12, 50);
            this.grpMargins.Name = "grpMargins";
            this.grpMargins.Size = new System.Drawing.Size(360, 120);
            this.grpMargins.TabIndex = 0;
            this.grpMargins.TabStop = false;
            this.grpMargins.Text = "边距设置 (像素)";
            //
            // btnSetAllMargins
            //
            this.btnSetAllMargins.Location = new System.Drawing.Point(270, 85);
            this.btnSetAllMargins.Name = "btnSetAllMargins";
            this.btnSetAllMargins.Size = new System.Drawing.Size(75, 25);
            this.btnSetAllMargins.TabIndex = 8;
            this.btnSetAllMargins.Text = "统一边距";
            this.btnSetAllMargins.UseVisualStyleBackColor = true;
            //
            // numRightMargin
            //
            this.numRightMargin.Location = new System.Drawing.Point(270, 55);
            this.numRightMargin.Maximum = new decimal(new int[] { 500, 0, 0, 0 });
            this.numRightMargin.Name = "numRightMargin";
            this.numRightMargin.Size = new System.Drawing.Size(75, 20);
            this.numRightMargin.TabIndex = 7;
            //
            // numLeftMargin
            //
            this.numLeftMargin.Location = new System.Drawing.Point(270, 25);
            this.numLeftMargin.Maximum = new decimal(new int[] { 500, 0, 0, 0 });
            this.numLeftMargin.Name = "numLeftMargin";
            this.numLeftMargin.Size = new System.Drawing.Size(75, 20);
            this.numLeftMargin.TabIndex = 6;
            //
            // numBottomMargin
            //
            this.numBottomMargin.Location = new System.Drawing.Point(80, 55);
            this.numBottomMargin.Maximum = new decimal(new int[] { 500, 0, 0, 0 });
            this.numBottomMargin.Name = "numBottomMargin";
            this.numBottomMargin.Size = new System.Drawing.Size(75, 20);
            this.numBottomMargin.TabIndex = 5;
            //
            // numTopMargin
            //
            this.numTopMargin.Location = new System.Drawing.Point(80, 25);
            this.numTopMargin.Maximum = new decimal(new int[] { 500, 0, 0, 0 });
            this.numTopMargin.Name = "numTopMargin";
            this.numTopMargin.Size = new System.Drawing.Size(75, 20);
            this.numTopMargin.TabIndex = 4;
            //
            // lblRightMargin
            //
            this.lblRightMargin.AutoSize = true;
            this.lblRightMargin.Location = new System.Drawing.Point(200, 57);
            this.lblRightMargin.Name = "lblRightMargin";
            this.lblRightMargin.Size = new System.Drawing.Size(55, 13);
            this.lblRightMargin.TabIndex = 3;
            this.lblRightMargin.Text = "右边距：";
            //
            // lblLeftMargin
            //
            this.lblLeftMargin.AutoSize = true;
            this.lblLeftMargin.Location = new System.Drawing.Point(200, 27);
            this.lblLeftMargin.Name = "lblLeftMargin";
            this.lblLeftMargin.Size = new System.Drawing.Size(55, 13);
            this.lblLeftMargin.TabIndex = 2;
            this.lblLeftMargin.Text = "左边距：";
            //
            // lblBottomMargin
            //
            this.lblBottomMargin.AutoSize = true;
            this.lblBottomMargin.Location = new System.Drawing.Point(10, 57);
            this.lblBottomMargin.Name = "lblBottomMargin";
            this.lblBottomMargin.Size = new System.Drawing.Size(55, 13);
            this.lblBottomMargin.TabIndex = 1;
            this.lblBottomMargin.Text = "下边距：";
            //
            // lblTopMargin
            //
            this.lblTopMargin.AutoSize = true;
            this.lblTopMargin.Location = new System.Drawing.Point(10, 27);
            this.lblTopMargin.Name = "lblTopMargin";
            this.lblTopMargin.Size = new System.Drawing.Size(55, 13);
            this.lblTopMargin.TabIndex = 0;
            this.lblTopMargin.Text = "上边距：";
            //
            // grpDetection
            //
            this.grpDetection.Controls.Add(this.lblMarginHint);
            this.grpDetection.Controls.Add(this.panelColorPreview);
            this.grpDetection.Controls.Add(this.btnCustomColor);
            this.grpDetection.Controls.Add(this.cmbBackgroundMode);
            this.grpDetection.Controls.Add(this.lblBackgroundMode);
            this.grpDetection.Controls.Add(this.numTolerance);
            this.grpDetection.Controls.Add(this.lblTolerance);
            this.grpDetection.Controls.Add(this.chkAutoDetect);
            this.grpDetection.Location = new System.Drawing.Point(12, 180);
            this.grpDetection.Name = "grpDetection";
            this.grpDetection.Size = new System.Drawing.Size(360, 140);
            this.grpDetection.TabIndex = 1;
            this.grpDetection.TabStop = false;
            this.grpDetection.Text = "自动检测设置";
            //
            // lblMarginHint
            //
            this.lblMarginHint.ForeColor = System.Drawing.Color.Gray;
            this.lblMarginHint.Location = new System.Drawing.Point(10, 115);
            this.lblMarginHint.Name = "lblMarginHint";
            this.lblMarginHint.Size = new System.Drawing.Size(340, 20);
            this.lblMarginHint.TabIndex = 7;
            this.lblMarginHint.Text = "边距将作为检测到的内容边界的额外缓冲区";
            //
            // panelColorPreview
            //
            this.panelColorPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelColorPreview.Location = new System.Drawing.Point(310, 85);
            this.panelColorPreview.Name = "panelColorPreview";
            this.panelColorPreview.Size = new System.Drawing.Size(25, 20);
            this.panelColorPreview.TabIndex = 6;
            //
            // btnCustomColor
            //
            this.btnCustomColor.Location = new System.Drawing.Point(200, 85);
            this.btnCustomColor.Name = "btnCustomColor";
            this.btnCustomColor.Size = new System.Drawing.Size(100, 23);
            this.btnCustomColor.TabIndex = 5;
            this.btnCustomColor.Text = "选择自定义颜色";
            this.btnCustomColor.UseVisualStyleBackColor = true;
            this.btnCustomColor.Click += new System.EventHandler(this.BtnCustomColor_Click);
            //
            // cmbBackgroundMode
            //
            this.cmbBackgroundMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbBackgroundMode.FormattingEnabled = true;
            this.cmbBackgroundMode.Location = new System.Drawing.Point(80, 85);
            this.cmbBackgroundMode.Name = "cmbBackgroundMode";
            this.cmbBackgroundMode.Size = new System.Drawing.Size(110, 21);
            this.cmbBackgroundMode.TabIndex = 4;
            //
            // lblBackgroundMode
            //
            this.lblBackgroundMode.AutoSize = true;
            this.lblBackgroundMode.Location = new System.Drawing.Point(10, 88);
            this.lblBackgroundMode.Name = "lblBackgroundMode";
            this.lblBackgroundMode.Size = new System.Drawing.Size(67, 13);
            this.lblBackgroundMode.TabIndex = 3;
            this.lblBackgroundMode.Text = "背景模式：";
            //
            // numTolerance
            //
            this.numTolerance.Location = new System.Drawing.Point(80, 55);
            this.numTolerance.Maximum = new decimal(new int[] { 50, 0, 0, 0 });
            this.numTolerance.Name = "numTolerance";
            this.numTolerance.Size = new System.Drawing.Size(75, 20);
            this.numTolerance.TabIndex = 2;
            //
            // lblTolerance
            //
            this.lblTolerance.AutoSize = true;
            this.lblTolerance.Location = new System.Drawing.Point(10, 57);
            this.lblTolerance.Name = "lblTolerance";
            this.lblTolerance.Size = new System.Drawing.Size(67, 13);
            this.lblTolerance.TabIndex = 1;
            this.lblTolerance.Text = "检测容差：";
            //
            // chkAutoDetect
            //
            this.chkAutoDetect.AutoSize = true;
            this.chkAutoDetect.Location = new System.Drawing.Point(10, 25);
            this.chkAutoDetect.Name = "chkAutoDetect";
            this.chkAutoDetect.Size = new System.Drawing.Size(98, 17);
            this.chkAutoDetect.TabIndex = 0;
            this.chkAutoDetect.Text = "自动检测边界";
            this.chkAutoDetect.UseVisualStyleBackColor = true;
            //
            // grpExport
            //
            this.grpExport.Controls.Add(this.cmbDpi);
            this.grpExport.Controls.Add(this.lblDpi);
            this.grpExport.Controls.Add(this.chkPreserveAspectRatio);
            this.grpExport.Controls.Add(this.numPdfQuality);
            this.grpExport.Controls.Add(this.lblPdfQuality);
            this.grpExport.Location = new System.Drawing.Point(12, 330);
            this.grpExport.Name = "grpExport";
            this.grpExport.Size = new System.Drawing.Size(360, 90);
            this.grpExport.TabIndex = 2;
            this.grpExport.TabStop = false;
            this.grpExport.Text = "导出设置";
            //
            // cmbDpi
            //
            this.cmbDpi.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbDpi.FormattingEnabled = true;
            this.cmbDpi.Location = new System.Drawing.Point(250, 25);
            this.cmbDpi.Name = "cmbDpi";
            this.cmbDpi.Size = new System.Drawing.Size(100, 21);
            this.cmbDpi.TabIndex = 4;
            //
            // lblDpi
            //
            this.lblDpi.AutoSize = true;
            this.lblDpi.Location = new System.Drawing.Point(200, 28);
            this.lblDpi.Name = "lblDpi";
            this.lblDpi.Size = new System.Drawing.Size(31, 13);
            this.lblDpi.TabIndex = 3;
            this.lblDpi.Text = "DPI：";
            //
            // chkPreserveAspectRatio
            //
            this.chkPreserveAspectRatio.AutoSize = true;
            this.chkPreserveAspectRatio.Location = new System.Drawing.Point(10, 60);
            this.chkPreserveAspectRatio.Name = "chkPreserveAspectRatio";
            this.chkPreserveAspectRatio.Size = new System.Drawing.Size(86, 17);
            this.chkPreserveAspectRatio.TabIndex = 2;
            this.chkPreserveAspectRatio.Text = "保持宽高比";
            this.chkPreserveAspectRatio.UseVisualStyleBackColor = true;
            //
            // numPdfQuality
            //
            this.numPdfQuality.Location = new System.Drawing.Point(80, 25);
            this.numPdfQuality.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            this.numPdfQuality.Name = "numPdfQuality";
            this.numPdfQuality.Size = new System.Drawing.Size(75, 20);
            this.numPdfQuality.TabIndex = 1;
            this.numPdfQuality.Value = new decimal(new int[] { 90, 0, 0, 0 });
            //
            // lblPdfQuality
            //
            this.lblPdfQuality.AutoSize = true;
            this.lblPdfQuality.Location = new System.Drawing.Point(10, 27);
            this.lblPdfQuality.Name = "lblPdfQuality";
            this.lblPdfQuality.Size = new System.Drawing.Size(67, 13);
            this.lblPdfQuality.TabIndex = 0;
            this.lblPdfQuality.Text = "PDF质量：";
            //
            // grpActions
            //
            this.grpActions.Controls.Add(this.btnExport);
            this.grpActions.Controls.Add(this.btnPreview);
            this.grpActions.Controls.Add(this.btnReset);
            this.grpActions.Controls.Add(this.btnCancel);
            this.grpActions.Controls.Add(this.btnSave);
            this.grpActions.Location = new System.Drawing.Point(12, 430);
            this.grpActions.Name = "grpActions";
            this.grpActions.Size = new System.Drawing.Size(360, 60);
            this.grpActions.TabIndex = 3;
            this.grpActions.TabStop = false;
            this.grpActions.Text = "操作";
            //
            // btnExport
            //
            this.btnExport.Location = new System.Drawing.Point(290, 25);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(60, 25);
            this.btnExport.TabIndex = 4;
            this.btnExport.Text = "导出PDF";
            this.btnExport.UseVisualStyleBackColor = true;
            //
            // btnPreview
            //
            this.btnPreview.Location = new System.Drawing.Point(220, 25);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(60, 25);
            this.btnPreview.TabIndex = 3;
            this.btnPreview.Text = "预览";
            this.btnPreview.UseVisualStyleBackColor = true;
            //
            // btnReset
            //
            this.btnReset.Location = new System.Drawing.Point(150, 25);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(60, 25);
            this.btnReset.TabIndex = 2;
            this.btnReset.Text = "重置";
            this.btnReset.UseVisualStyleBackColor = true;
            //
            // btnCancel
            //
            this.btnCancel.Location = new System.Drawing.Point(80, 25);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(60, 25);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = true;
            //
            // btnSave
            //
            this.btnSave.Location = new System.Drawing.Point(10, 25);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(60, 25);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            //
            // lblValidation
            //
            this.lblValidation.Location = new System.Drawing.Point(12, 500);
            this.lblValidation.Name = "lblValidation";
            this.lblValidation.Size = new System.Drawing.Size(360, 20);
            this.lblValidation.TabIndex = 4;
            this.lblValidation.Text = "设置有效";
            this.lblValidation.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            //
            // lblTitle
            //
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(12, 15);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(228, 20);
            this.lblTitle.TabIndex = 5;
            this.lblTitle.Text = "FreeCut - PPT自动裁剪PDF导出";
            //
            // SettingsForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 530);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.lblValidation);
            this.Controls.Add(this.grpActions);
            this.Controls.Add(this.grpExport);
            this.Controls.Add(this.grpDetection);
            this.Controls.Add(this.grpMargins);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FreeCut 设置";
            this.grpMargins.ResumeLayout(false);
            this.grpMargins.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numRightMargin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLeftMargin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numBottomMargin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTopMargin)).EndInit();
            this.grpDetection.ResumeLayout(false);
            this.grpDetection.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numTolerance)).EndInit();
            this.grpExport.ResumeLayout(false);
            this.grpExport.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numPdfQuality)).EndInit();
            this.grpActions.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox grpMargins;
        private System.Windows.Forms.Button btnSetAllMargins;
        private System.Windows.Forms.NumericUpDown numRightMargin;
        private System.Windows.Forms.NumericUpDown numLeftMargin;
        private System.Windows.Forms.NumericUpDown numBottomMargin;
        private System.Windows.Forms.NumericUpDown numTopMargin;
        private System.Windows.Forms.Label lblRightMargin;
        private System.Windows.Forms.Label lblLeftMargin;
        private System.Windows.Forms.Label lblBottomMargin;
        private System.Windows.Forms.Label lblTopMargin;
        private System.Windows.Forms.GroupBox grpDetection;
        private System.Windows.Forms.Label lblMarginHint;
        private System.Windows.Forms.Panel panelColorPreview;
        private System.Windows.Forms.Button btnCustomColor;
        private System.Windows.Forms.ComboBox cmbBackgroundMode;
        private System.Windows.Forms.Label lblBackgroundMode;
        private System.Windows.Forms.NumericUpDown numTolerance;
        private System.Windows.Forms.Label lblTolerance;
        private System.Windows.Forms.CheckBox chkAutoDetect;
        private System.Windows.Forms.GroupBox grpExport;
        private System.Windows.Forms.ComboBox cmbDpi;
        private System.Windows.Forms.Label lblDpi;
        private System.Windows.Forms.CheckBox chkPreserveAspectRatio;
        private System.Windows.Forms.NumericUpDown numPdfQuality;
        private System.Windows.Forms.Label lblPdfQuality;
        private System.Windows.Forms.GroupBox grpActions;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnPreview;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label lblValidation;
        private System.Windows.Forms.Label lblTitle;
    }
}