namespace ReporterLoaders
{
    partial class MainForm
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.txtStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.btnOpen = new System.Windows.Forms.ToolStripButton();
            this.btnExportTo = new System.Windows.Forms.ToolStripButton();
            this.btnUploadData = new System.Windows.Forms.ToolStripButton();
            this.btnConfig = new System.Windows.Forms.ToolStripButton();
            this.btnExit = new System.Windows.Forms.ToolStripButton();
            this.plMain = new System.Windows.Forms.Panel();
            this.dgvDetail = new System.Windows.Forms.DataGridView();
            this.colHead = new System.Windows.Forms.DataGridViewImageColumn();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSex = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMinZu = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colZhengzhiMianMao = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colZhuanYe = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colXueWei = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colWork = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colUnit = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colJob = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSheMiChengDu = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colTelphone = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colButtons = new System.Windows.Forms.DataGridViewButtonColumn();
            this.ofdWords = new System.Windows.Forms.OpenFileDialog();
            this.statusStrip.SuspendLayout();
            this.toolStrip.SuspendLayout();
            this.plMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDetail)).BeginInit();
            this.SuspendLayout();
            // 
            // statusStrip
            // 
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.txtStatus});
            this.statusStrip.Location = new System.Drawing.Point(0, 532);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(1080, 22);
            this.statusStrip.TabIndex = 0;
            this.statusStrip.Text = "statusStrip1";
            // 
            // txtStatus
            // 
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.Size = new System.Drawing.Size(47, 17);
            this.txtStatus.Text = "Welcome";
            // 
            // toolStrip
            // 
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnOpen,
            this.btnExportTo,
            this.btnUploadData,
            this.btnConfig,
            this.btnExit});
            this.toolStrip.Location = new System.Drawing.Point(0, 0);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.Size = new System.Drawing.Size(1080, 91);
            this.toolStrip.TabIndex = 1;
            this.toolStrip.Text = "toolStrip1";
            // 
            // btnOpen
            // 
            this.btnOpen.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnOpen.Image = ((System.Drawing.Image)(resources.GetObject("btnOpen.Image")));
            this.btnOpen.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.btnOpen.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(69, 88);
            this.btnOpen.Text = "批量载入";
            this.btnOpen.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // btnExportTo
            // 
            this.btnExportTo.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExportTo.Image = ((System.Drawing.Image)(resources.GetObject("btnExportTo.Image")));
            this.btnExportTo.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.btnExportTo.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnExportTo.Name = "btnExportTo";
            this.btnExportTo.Size = new System.Drawing.Size(89, 88);
            this.btnExportTo.Text = "导出到Excel";
            this.btnExportTo.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnExportTo.Click += new System.EventHandler(this.btnExportTo_Click);
            // 
            // btnUploadData
            // 
            this.btnUploadData.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnUploadData.Image = ((System.Drawing.Image)(resources.GetObject("btnUploadData.Image")));
            this.btnUploadData.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.btnUploadData.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnUploadData.Name = "btnUploadData";
            this.btnUploadData.Size = new System.Drawing.Size(69, 88);
            this.btnUploadData.Text = "数据上传";
            this.btnUploadData.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnUploadData.Click += new System.EventHandler(this.btnUploadData_Click);
            // 
            // btnConfig
            // 
            this.btnConfig.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnConfig.Image = ((System.Drawing.Image)(resources.GetObject("btnConfig.Image")));
            this.btnConfig.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.btnConfig.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnConfig.Name = "btnConfig";
            this.btnConfig.Size = new System.Drawing.Size(69, 88);
            this.btnConfig.Text = "系统配置";
            this.btnConfig.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnConfig.Click += new System.EventHandler(this.btnConfig_Click);
            // 
            // btnExit
            // 
            this.btnExit.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExit.Image = ((System.Drawing.Image)(resources.GetObject("btnExit.Image")));
            this.btnExit.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None;
            this.btnExit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(68, 88);
            this.btnExit.Text = "退出";
            this.btnExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // plMain
            // 
            this.plMain.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.plMain.Controls.Add(this.dgvDetail);
            this.plMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.plMain.Location = new System.Drawing.Point(0, 91);
            this.plMain.Name = "plMain";
            this.plMain.Size = new System.Drawing.Size(1080, 441);
            this.plMain.TabIndex = 2;
            // 
            // dgvDetail
            // 
            this.dgvDetail.AllowUserToAddRows = false;
            this.dgvDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDetail.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colHead,
            this.colName,
            this.colSex,
            this.colMinZu,
            this.colZhengzhiMianMao,
            this.colZhuanYe,
            this.colXueWei,
            this.colCity,
            this.colWork,
            this.colUnit,
            this.colJob,
            this.colSheMiChengDu,
            this.colTelphone,
            this.colButtons});
            this.dgvDetail.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvDetail.Location = new System.Drawing.Point(0, 0);
            this.dgvDetail.Margin = new System.Windows.Forms.Padding(0);
            this.dgvDetail.Name = "dgvDetail";
            this.dgvDetail.ReadOnly = true;
            this.dgvDetail.RowTemplate.Height = 128;
            this.dgvDetail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvDetail.Size = new System.Drawing.Size(1080, 441);
            this.dgvDetail.TabIndex = 0;
            this.dgvDetail.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvDetail_CellContentClick);
            // 
            // colHead
            // 
            this.colHead.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.colHead.HeaderText = "";
            this.colHead.Name = "colHead";
            this.colHead.ReadOnly = true;
            this.colHead.Width = 5;
            // 
            // colName
            // 
            this.colName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colName.HeaderText = "姓名";
            this.colName.Name = "colName";
            this.colName.ReadOnly = true;
            this.colName.Width = 54;
            // 
            // colSex
            // 
            this.colSex.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colSex.HeaderText = "性别";
            this.colSex.Name = "colSex";
            this.colSex.ReadOnly = true;
            this.colSex.Width = 54;
            // 
            // colMinZu
            // 
            this.colMinZu.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colMinZu.HeaderText = "民族";
            this.colMinZu.Name = "colMinZu";
            this.colMinZu.ReadOnly = true;
            this.colMinZu.Width = 54;
            // 
            // colZhengzhiMianMao
            // 
            this.colZhengzhiMianMao.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colZhengzhiMianMao.HeaderText = "政治面貌";
            this.colZhengzhiMianMao.Name = "colZhengzhiMianMao";
            this.colZhengzhiMianMao.ReadOnly = true;
            this.colZhengzhiMianMao.Width = 78;
            // 
            // colZhuanYe
            // 
            this.colZhuanYe.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colZhuanYe.HeaderText = "专业";
            this.colZhuanYe.Name = "colZhuanYe";
            this.colZhuanYe.ReadOnly = true;
            this.colZhuanYe.Width = 54;
            // 
            // colXueWei
            // 
            this.colXueWei.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colXueWei.HeaderText = "学位";
            this.colXueWei.Name = "colXueWei";
            this.colXueWei.ReadOnly = true;
            this.colXueWei.Width = 54;
            // 
            // colCity
            // 
            this.colCity.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colCity.HeaderText = "所在省市";
            this.colCity.Name = "colCity";
            this.colCity.ReadOnly = true;
            this.colCity.Width = 78;
            // 
            // colWork
            // 
            this.colWork.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colWork.HeaderText = "工作单位";
            this.colWork.Name = "colWork";
            this.colWork.ReadOnly = true;
            this.colWork.Width = 78;
            // 
            // colUnit
            // 
            this.colUnit.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colUnit.HeaderText = "所属部门";
            this.colUnit.Name = "colUnit";
            this.colUnit.ReadOnly = true;
            this.colUnit.Width = 78;
            // 
            // colJob
            // 
            this.colJob.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colJob.HeaderText = "行政职务";
            this.colJob.Name = "colJob";
            this.colJob.ReadOnly = true;
            this.colJob.Width = 78;
            // 
            // colSheMiChengDu
            // 
            this.colSheMiChengDu.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colSheMiChengDu.HeaderText = "涉密程度";
            this.colSheMiChengDu.Name = "colSheMiChengDu";
            this.colSheMiChengDu.ReadOnly = true;
            this.colSheMiChengDu.Width = 78;
            // 
            // colTelphone
            // 
            this.colTelphone.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colTelphone.HeaderText = "电话";
            this.colTelphone.Name = "colTelphone";
            this.colTelphone.ReadOnly = true;
            this.colTelphone.Width = 54;
            // 
            // colButtons
            // 
            this.colButtons.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.colButtons.HeaderText = "";
            this.colButtons.Name = "colButtons";
            this.colButtons.ReadOnly = true;
            this.colButtons.Text = "删除";
            this.colButtons.UseColumnTextForButtonValue = true;
            this.colButtons.Width = 70;
            // 
            // ofdWords
            // 
            this.ofdWords.Filter = "*.docx|*.docx|*.doc|*.doc";
            this.ofdWords.Multiselect = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1080, 554);
            this.Controls.Add(this.plMain);
            this.Controls.Add(this.toolStrip);
            this.Controls.Add(this.statusStrip);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "人员信息导入";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.toolStrip.ResumeLayout(false);
            this.toolStrip.PerformLayout();
            this.plMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDetail)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel txtStatus;
        private System.Windows.Forms.ToolStrip toolStrip;
        private System.Windows.Forms.Panel plMain;
        private System.Windows.Forms.ToolStripButton btnOpen;
        private System.Windows.Forms.ToolStripButton btnExit;
        private System.Windows.Forms.ToolStripButton btnConfig;
        private System.Windows.Forms.ToolStripButton btnUploadData;
        private System.Windows.Forms.ToolStripButton btnExportTo;
        private System.Windows.Forms.DataGridView dgvDetail;
        private System.Windows.Forms.OpenFileDialog ofdWords;
        private System.Windows.Forms.DataGridViewImageColumn colHead;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSex;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMinZu;
        private System.Windows.Forms.DataGridViewTextBoxColumn colZhengzhiMianMao;
        private System.Windows.Forms.DataGridViewTextBoxColumn colZhuanYe;
        private System.Windows.Forms.DataGridViewTextBoxColumn colXueWei;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCity;
        private System.Windows.Forms.DataGridViewTextBoxColumn colWork;
        private System.Windows.Forms.DataGridViewTextBoxColumn colUnit;
        private System.Windows.Forms.DataGridViewTextBoxColumn colJob;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSheMiChengDu;
        private System.Windows.Forms.DataGridViewTextBoxColumn colTelphone;
        private System.Windows.Forms.DataGridViewButtonColumn colButtons;
    }
}

