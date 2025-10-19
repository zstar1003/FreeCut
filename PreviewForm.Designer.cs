namespace FreeCut
{
    partial class PreviewForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.pictureBoxPreview = new System.Windows.Forms.PictureBox();
            this.btnPrevious = new System.Windows.Forms.Button();
            this.btnNext = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.lblSlideInfo = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.panelNavigation = new System.Windows.Forms.Panel();
            this.panelPreview = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxPreview)).BeginInit();
            this.panelNavigation.SuspendLayout();
            this.panelPreview.SuspendLayout();
            this.SuspendLayout();
            //
            // pictureBoxPreview
            //
            this.pictureBoxPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBoxPreview.BackColor = System.Drawing.Color.White;
            this.pictureBoxPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pictureBoxPreview.Location = new System.Drawing.Point(10, 10);
            this.pictureBoxPreview.Name = "pictureBoxPreview";
            this.pictureBoxPreview.Size = new System.Drawing.Size(660, 460);
            this.pictureBoxPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBoxPreview.TabIndex = 0;
            this.pictureBoxPreview.TabStop = false;
            //
            // btnPrevious
            //
            this.btnPrevious.Location = new System.Drawing.Point(10, 10);
            this.btnPrevious.Name = "btnPrevious";
            this.btnPrevious.Size = new System.Drawing.Size(75, 30);
            this.btnPrevious.TabIndex = 1;
            this.btnPrevious.Text = "← 上一张";
            this.btnPrevious.UseVisualStyleBackColor = true;
            //
            // btnNext
            //
            this.btnNext.Location = new System.Drawing.Point(95, 10);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(75, 30);
            this.btnNext.TabIndex = 2;
            this.btnNext.Text = "下一张 →";
            this.btnNext.UseVisualStyleBackColor = true;
            //
            // btnClose
            //
            this.btnClose.Location = new System.Drawing.Point(595, 10);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 30);
            this.btnClose.TabIndex = 3;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = true;
            //
            // btnExport
            //
            this.btnExport.Location = new System.Drawing.Point(510, 10);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 30);
            this.btnExport.TabIndex = 4;
            this.btnExport.Text = "导出PDF";
            this.btnExport.UseVisualStyleBackColor = true;
            //
            // btnRefresh
            //
            this.btnRefresh.Location = new System.Drawing.Point(425, 10);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(75, 30);
            this.btnRefresh.TabIndex = 5;
            this.btnRefresh.Text = "刷新预览";
            this.btnRefresh.UseVisualStyleBackColor = true;
            //
            // lblSlideInfo
            //
            this.lblSlideInfo.AutoSize = true;
            this.lblSlideInfo.Location = new System.Drawing.Point(180, 20);
            this.lblSlideInfo.Name = "lblSlideInfo";
            this.lblSlideInfo.Size = new System.Drawing.Size(65, 13);
            this.lblSlideInfo.TabIndex = 6;
            this.lblSlideInfo.Text = "第 1 / 1 张";
            //
            // lblStatus
            //
            this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblStatus.ForeColor = System.Drawing.Color.Gray;
            this.lblStatus.Location = new System.Drawing.Point(10, 540);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(680, 15);
            this.lblStatus.TabIndex = 7;
            this.lblStatus.Text = "就绪";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // lblTitle
            //
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(12, 15);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(144, 17);
            this.lblTitle.TabIndex = 8;
            this.lblTitle.Text = "FreeCut 裁剪预览";
            //
            // panelNavigation
            //
            this.panelNavigation.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelNavigation.Controls.Add(this.btnPrevious);
            this.panelNavigation.Controls.Add(this.btnNext);
            this.panelNavigation.Controls.Add(this.btnRefresh);
            this.panelNavigation.Controls.Add(this.btnExport);
            this.panelNavigation.Controls.Add(this.btnClose);
            this.panelNavigation.Controls.Add(this.lblSlideInfo);
            this.panelNavigation.Location = new System.Drawing.Point(10, 485);
            this.panelNavigation.Name = "panelNavigation";
            this.panelNavigation.Size = new System.Drawing.Size(680, 50);
            this.panelNavigation.TabIndex = 9;
            //
            // panelPreview
            //
            this.panelPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelPreview.Controls.Add(this.pictureBoxPreview);
            this.panelPreview.Location = new System.Drawing.Point(10, 40);
            this.panelPreview.Name = "panelPreview";
            this.panelPreview.Size = new System.Drawing.Size(680, 440);
            this.panelPreview.TabIndex = 10;
            //
            // PreviewForm
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(700, 565);
            this.Controls.Add(this.panelPreview);
            this.Controls.Add(this.panelNavigation);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.lblStatus);
            this.MinimumSize = new System.Drawing.Size(600, 400);
            this.Name = "PreviewForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FreeCut 预览";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxPreview)).EndInit();
            this.panelNavigation.ResumeLayout(false);
            this.panelNavigation.PerformLayout();
            this.panelPreview.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBoxPreview;
        private System.Windows.Forms.Button btnPrevious;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Label lblSlideInfo;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Panel panelNavigation;
        private System.Windows.Forms.Panel panelPreview;
    }
}