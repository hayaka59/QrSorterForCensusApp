namespace QrSorterInspectionApp
{
    partial class LogListForm
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
            this.LblVersion = new System.Windows.Forms.Label();
            this.LblTitle = new System.Windows.Forms.Label();
            this.CmbLogType = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.LsvLogContent = new System.Windows.Forms.ListView();
            this.label9 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.CmbReasonForNonDelivery1 = new System.Windows.Forms.ComboBox();
            this.dtTimePickerFrom = new System.Windows.Forms.DateTimePicker();
            this.ChkInspectionDate = new System.Windows.Forms.CheckBox();
            this.GrpInspectionDate = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.dtTimePickerTo = new System.Windows.Forms.DateTimePicker();
            this.GrpReasonForNonDelivery1 = new System.Windows.Forms.GroupBox();
            this.ChkReasonForNonDelivery1 = new System.Windows.Forms.CheckBox();
            this.LblLogFileCount = new System.Windows.Forms.Label();
            this.LblContentCount = new System.Windows.Forms.Label();
            this.LblSelectedFile = new System.Windows.Forms.Label();
            this.BtnJobClear = new System.Windows.Forms.Button();
            this.GrpReasonForNonDelivery2 = new System.Windows.Forms.GroupBox();
            this.ChkReasonForNonDelivery2 = new System.Windows.Forms.CheckBox();
            this.CmbReasonForNonDelivery2 = new System.Windows.Forms.ComboBox();
            this.GrpSortBy = new System.Windows.Forms.GroupBox();
            this.CmbSortBy = new System.Windows.Forms.ComboBox();
            this.PicWaitList = new System.Windows.Forms.PictureBox();
            this.BtnJobSelect = new System.Windows.Forms.Button();
            this.BtnUpdate = new System.Windows.Forms.Button();
            this.BtnClose = new System.Windows.Forms.Button();
            this.PicWaitContent = new System.Windows.Forms.PictureBox();
            this.LsvLogList = new System.Windows.Forms.ListView();
            this.GrpInspectionDate.SuspendLayout();
            this.GrpReasonForNonDelivery1.SuspendLayout();
            this.GrpReasonForNonDelivery2.SuspendLayout();
            this.GrpSortBy.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.PicWaitList)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PicWaitContent)).BeginInit();
            this.SuspendLayout();
            // 
            // LblVersion
            // 
            this.LblVersion.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblVersion.ForeColor = System.Drawing.Color.Blue;
            this.LblVersion.Location = new System.Drawing.Point(1758, 998);
            this.LblVersion.Name = "LblVersion";
            this.LblVersion.Size = new System.Drawing.Size(130, 25);
            this.LblVersion.TabIndex = 242;
            this.LblVersion.Text = "LblVersion";
            this.LblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // LblTitle
            // 
            this.LblTitle.BackColor = System.Drawing.Color.CornflowerBlue;
            this.LblTitle.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblTitle.ForeColor = System.Drawing.Color.White;
            this.LblTitle.Location = new System.Drawing.Point(1, 1);
            this.LblTitle.Name = "LblTitle";
            this.LblTitle.Size = new System.Drawing.Size(1909, 50);
            this.LblTitle.TabIndex = 240;
            this.LblTitle.Text = "ログ管理";
            this.LblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CmbLogType
            // 
            this.CmbLogType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbLogType.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CmbLogType.FormattingEnabled = true;
            this.CmbLogType.IntegralHeight = false;
            this.CmbLogType.ItemHeight = 28;
            this.CmbLogType.Location = new System.Drawing.Point(171, 88);
            this.CmbLogType.Name = "CmbLogType";
            this.CmbLogType.Size = new System.Drawing.Size(134, 36);
            this.CmbLogType.TabIndex = 261;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(38, 87);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(132, 36);
            this.label3.TabIndex = 260;
            this.label3.Text = "ログの種類";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // LsvLogContent
            // 
            this.LsvLogContent.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LsvLogContent.FullRowSelect = true;
            this.LsvLogContent.GridLines = true;
            this.LsvLogContent.HideSelection = false;
            this.LsvLogContent.Location = new System.Drawing.Point(38, 452);
            this.LsvLogContent.Name = "LsvLogContent";
            this.LsvLogContent.Size = new System.Drawing.Size(1827, 488);
            this.LsvLogContent.TabIndex = 258;
            this.LsvLogContent.UseCompatibleStateImageBehavior = false;
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.Color.CornflowerBlue;
            this.label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label9.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label9.ForeColor = System.Drawing.Color.White;
            this.label9.Location = new System.Drawing.Point(38, 422);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(1827, 31);
            this.label9.TabIndex = 259;
            this.label9.Text = "選択したログの内容";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.CornflowerBlue;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(38, 152);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(1827, 31);
            this.label1.TabIndex = 262;
            this.label1.Text = "検査ログ一覧";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CmbReasonForNonDelivery1
            // 
            this.CmbReasonForNonDelivery1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbReasonForNonDelivery1.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CmbReasonForNonDelivery1.FormattingEnabled = true;
            this.CmbReasonForNonDelivery1.IntegralHeight = false;
            this.CmbReasonForNonDelivery1.ItemHeight = 28;
            this.CmbReasonForNonDelivery1.Location = new System.Drawing.Point(15, 38);
            this.CmbReasonForNonDelivery1.Name = "CmbReasonForNonDelivery1";
            this.CmbReasonForNonDelivery1.Size = new System.Drawing.Size(207, 36);
            this.CmbReasonForNonDelivery1.TabIndex = 265;
            // 
            // dtTimePickerFrom
            // 
            this.dtTimePickerFrom.CalendarFont = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dtTimePickerFrom.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dtTimePickerFrom.Location = new System.Drawing.Point(15, 41);
            this.dtTimePickerFrom.Name = "dtTimePickerFrom";
            this.dtTimePickerFrom.Size = new System.Drawing.Size(171, 31);
            this.dtTimePickerFrom.TabIndex = 266;
            // 
            // ChkInspectionDate
            // 
            this.ChkInspectionDate.AutoSize = true;
            this.ChkInspectionDate.Location = new System.Drawing.Point(17, 20);
            this.ChkInspectionDate.Name = "ChkInspectionDate";
            this.ChkInspectionDate.Size = new System.Drawing.Size(103, 16);
            this.ChkInspectionDate.TabIndex = 267;
            this.ChkInspectionDate.Text = "更新条件に含む";
            this.ChkInspectionDate.UseVisualStyleBackColor = true;
            // 
            // GrpInspectionDate
            // 
            this.GrpInspectionDate.Controls.Add(this.label4);
            this.GrpInspectionDate.Controls.Add(this.dtTimePickerTo);
            this.GrpInspectionDate.Controls.Add(this.ChkInspectionDate);
            this.GrpInspectionDate.Controls.Add(this.dtTimePickerFrom);
            this.GrpInspectionDate.Location = new System.Drawing.Point(794, 58);
            this.GrpInspectionDate.Name = "GrpInspectionDate";
            this.GrpInspectionDate.Size = new System.Drawing.Size(407, 85);
            this.GrpInspectionDate.TabIndex = 268;
            this.GrpInspectionDate.TabStop = false;
            this.GrpInspectionDate.Text = "検査日付";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.Location = new System.Drawing.Point(190, 45);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(26, 24);
            this.label4.TabIndex = 269;
            this.label4.Text = "～";
            // 
            // dtTimePickerTo
            // 
            this.dtTimePickerTo.CalendarFont = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dtTimePickerTo.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dtTimePickerTo.Location = new System.Drawing.Point(219, 41);
            this.dtTimePickerTo.Name = "dtTimePickerTo";
            this.dtTimePickerTo.Size = new System.Drawing.Size(171, 31);
            this.dtTimePickerTo.TabIndex = 268;
            // 
            // GrpReasonForNonDelivery1
            // 
            this.GrpReasonForNonDelivery1.Controls.Add(this.ChkReasonForNonDelivery1);
            this.GrpReasonForNonDelivery1.Controls.Add(this.CmbReasonForNonDelivery1);
            this.GrpReasonForNonDelivery1.Location = new System.Drawing.Point(1207, 58);
            this.GrpReasonForNonDelivery1.Name = "GrpReasonForNonDelivery1";
            this.GrpReasonForNonDelivery1.Size = new System.Drawing.Size(234, 85);
            this.GrpReasonForNonDelivery1.TabIndex = 270;
            this.GrpReasonForNonDelivery1.TabStop = false;
            this.GrpReasonForNonDelivery1.Text = "不着事由仕分け１";
            // 
            // ChkReasonForNonDelivery1
            // 
            this.ChkReasonForNonDelivery1.AutoSize = true;
            this.ChkReasonForNonDelivery1.Location = new System.Drawing.Point(17, 18);
            this.ChkReasonForNonDelivery1.Name = "ChkReasonForNonDelivery1";
            this.ChkReasonForNonDelivery1.Size = new System.Drawing.Size(103, 16);
            this.ChkReasonForNonDelivery1.TabIndex = 267;
            this.ChkReasonForNonDelivery1.Text = "更新条件に含む";
            this.ChkReasonForNonDelivery1.UseVisualStyleBackColor = true;
            // 
            // LblLogFileCount
            // 
            this.LblLogFileCount.BackColor = System.Drawing.Color.CornflowerBlue;
            this.LblLogFileCount.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblLogFileCount.ForeColor = System.Drawing.Color.White;
            this.LblLogFileCount.Location = new System.Drawing.Point(52, 157);
            this.LblLogFileCount.Name = "LblLogFileCount";
            this.LblLogFileCount.Size = new System.Drawing.Size(536, 23);
            this.LblLogFileCount.TabIndex = 271;
            this.LblLogFileCount.Text = "LblLogFileCount";
            this.LblLogFileCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // LblContentCount
            // 
            this.LblContentCount.BackColor = System.Drawing.Color.CornflowerBlue;
            this.LblContentCount.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblContentCount.ForeColor = System.Drawing.Color.White;
            this.LblContentCount.Location = new System.Drawing.Point(52, 427);
            this.LblContentCount.Name = "LblContentCount";
            this.LblContentCount.Size = new System.Drawing.Size(536, 23);
            this.LblContentCount.TabIndex = 272;
            this.LblContentCount.Text = "LblContentCount";
            this.LblContentCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // LblSelectedFile
            // 
            this.LblSelectedFile.BackColor = System.Drawing.Color.White;
            this.LblSelectedFile.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblSelectedFile.ForeColor = System.Drawing.Color.Black;
            this.LblSelectedFile.Location = new System.Drawing.Point(316, 107);
            this.LblSelectedFile.Name = "LblSelectedFile";
            this.LblSelectedFile.Size = new System.Drawing.Size(298, 36);
            this.LblSelectedFile.TabIndex = 325;
            this.LblSelectedFile.Text = "LblSelectedFile";
            this.LblSelectedFile.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // BtnJobClear
            // 
            this.BtnJobClear.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnJobClear.Location = new System.Drawing.Point(472, 58);
            this.BtnJobClear.Name = "BtnJobClear";
            this.BtnJobClear.Size = new System.Drawing.Size(144, 46);
            this.BtnJobClear.TabIndex = 326;
            this.BtnJobClear.Text = "JOBクリア";
            this.BtnJobClear.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnJobClear.UseVisualStyleBackColor = true;
            this.BtnJobClear.Click += new System.EventHandler(this.BtnJobClear_Click);
            // 
            // GrpReasonForNonDelivery2
            // 
            this.GrpReasonForNonDelivery2.Controls.Add(this.ChkReasonForNonDelivery2);
            this.GrpReasonForNonDelivery2.Controls.Add(this.CmbReasonForNonDelivery2);
            this.GrpReasonForNonDelivery2.Location = new System.Drawing.Point(1453, 58);
            this.GrpReasonForNonDelivery2.Name = "GrpReasonForNonDelivery2";
            this.GrpReasonForNonDelivery2.Size = new System.Drawing.Size(234, 85);
            this.GrpReasonForNonDelivery2.TabIndex = 271;
            this.GrpReasonForNonDelivery2.TabStop = false;
            this.GrpReasonForNonDelivery2.Text = "不着事由仕分け２";
            // 
            // ChkReasonForNonDelivery2
            // 
            this.ChkReasonForNonDelivery2.AutoSize = true;
            this.ChkReasonForNonDelivery2.Location = new System.Drawing.Point(17, 18);
            this.ChkReasonForNonDelivery2.Name = "ChkReasonForNonDelivery2";
            this.ChkReasonForNonDelivery2.Size = new System.Drawing.Size(103, 16);
            this.ChkReasonForNonDelivery2.TabIndex = 267;
            this.ChkReasonForNonDelivery2.Text = "更新条件に含む";
            this.ChkReasonForNonDelivery2.UseVisualStyleBackColor = true;
            // 
            // CmbReasonForNonDelivery2
            // 
            this.CmbReasonForNonDelivery2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbReasonForNonDelivery2.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CmbReasonForNonDelivery2.FormattingEnabled = true;
            this.CmbReasonForNonDelivery2.IntegralHeight = false;
            this.CmbReasonForNonDelivery2.ItemHeight = 28;
            this.CmbReasonForNonDelivery2.Location = new System.Drawing.Point(15, 38);
            this.CmbReasonForNonDelivery2.Name = "CmbReasonForNonDelivery2";
            this.CmbReasonForNonDelivery2.Size = new System.Drawing.Size(207, 36);
            this.CmbReasonForNonDelivery2.TabIndex = 265;
            // 
            // GrpSortBy
            // 
            this.GrpSortBy.Controls.Add(this.CmbSortBy);
            this.GrpSortBy.Location = new System.Drawing.Point(632, 74);
            this.GrpSortBy.Name = "GrpSortBy";
            this.GrpSortBy.Size = new System.Drawing.Size(146, 56);
            this.GrpSortBy.TabIndex = 272;
            this.GrpSortBy.TabStop = false;
            this.GrpSortBy.Text = "並び順";
            // 
            // CmbSortBy
            // 
            this.CmbSortBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbSortBy.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CmbSortBy.FormattingEnabled = true;
            this.CmbSortBy.IntegralHeight = false;
            this.CmbSortBy.ItemHeight = 23;
            this.CmbSortBy.Location = new System.Drawing.Point(7, 15);
            this.CmbSortBy.Name = "CmbSortBy";
            this.CmbSortBy.Size = new System.Drawing.Size(133, 31);
            this.CmbSortBy.TabIndex = 265;
            // 
            // PicWaitList
            // 
            this.PicWaitList.Image = global::QrSorterInspectionApp.Properties.Resources.waiting;
            this.PicWaitList.Location = new System.Drawing.Point(901, 237);
            this.PicWaitList.Name = "PicWaitList";
            this.PicWaitList.Size = new System.Drawing.Size(100, 100);
            this.PicWaitList.TabIndex = 327;
            this.PicWaitList.TabStop = false;
            this.PicWaitList.Visible = false;
            // 
            // BtnJobSelect
            // 
            this.BtnJobSelect.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnJobSelect.Image = global::QrSorterInspectionApp.Properties.Resources.search_file;
            this.BtnJobSelect.Location = new System.Drawing.Point(314, 58);
            this.BtnJobSelect.Name = "BtnJobSelect";
            this.BtnJobSelect.Size = new System.Drawing.Size(144, 46);
            this.BtnJobSelect.TabIndex = 324;
            this.BtnJobSelect.Text = "JOB選択";
            this.BtnJobSelect.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnJobSelect.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnJobSelect.UseVisualStyleBackColor = true;
            this.BtnJobSelect.Click += new System.EventHandler(this.BtnJobSelect_Click);
            // 
            // BtnUpdate
            // 
            this.BtnUpdate.BackColor = System.Drawing.Color.LightGreen;
            this.BtnUpdate.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnUpdate.Image = global::QrSorterInspectionApp.Properties.Resources.update;
            this.BtnUpdate.Location = new System.Drawing.Point(1697, 58);
            this.BtnUpdate.Name = "BtnUpdate";
            this.BtnUpdate.Size = new System.Drawing.Size(168, 85);
            this.BtnUpdate.TabIndex = 273;
            this.BtnUpdate.Text = "更新";
            this.BtnUpdate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnUpdate.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnUpdate.UseVisualStyleBackColor = false;
            this.BtnUpdate.Click += new System.EventHandler(this.BtnUpdate_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnClose.Image = global::QrSorterInspectionApp.Properties.Resources.back_arrow;
            this.BtnClose.Location = new System.Drawing.Point(1547, 973);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(205, 50);
            this.BtnClose.TabIndex = 241;
            this.BtnClose.Text = "戻る";
            this.BtnClose.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnClose.UseVisualStyleBackColor = true;
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // PicWaitContent
            // 
            this.PicWaitContent.Image = global::QrSorterInspectionApp.Properties.Resources.waiting;
            this.PicWaitContent.Location = new System.Drawing.Point(902, 602);
            this.PicWaitContent.Name = "PicWaitContent";
            this.PicWaitContent.Size = new System.Drawing.Size(100, 100);
            this.PicWaitContent.TabIndex = 328;
            this.PicWaitContent.TabStop = false;
            this.PicWaitContent.Visible = false;
            // 
            // LsvLogList
            // 
            this.LsvLogList.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LsvLogList.FullRowSelect = true;
            this.LsvLogList.GridLines = true;
            this.LsvLogList.HideSelection = false;
            this.LsvLogList.Location = new System.Drawing.Point(38, 181);
            this.LsvLogList.Name = "LsvLogList";
            this.LsvLogList.Size = new System.Drawing.Size(1827, 221);
            this.LsvLogList.TabIndex = 329;
            this.LsvLogList.UseCompatibleStateImageBehavior = false;
            this.LsvLogList.SelectedIndexChanged += new System.EventHandler(this.LsvLogList_SelectedIndexChanged);
            // 
            // LogListForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1904, 1066);
            this.ControlBox = false;
            this.Controls.Add(this.PicWaitContent);
            this.Controls.Add(this.PicWaitList);
            this.Controls.Add(this.GrpSortBy);
            this.Controls.Add(this.GrpReasonForNonDelivery2);
            this.Controls.Add(this.BtnJobClear);
            this.Controls.Add(this.LblSelectedFile);
            this.Controls.Add(this.BtnJobSelect);
            this.Controls.Add(this.BtnUpdate);
            this.Controls.Add(this.LblContentCount);
            this.Controls.Add(this.LblLogFileCount);
            this.Controls.Add(this.GrpReasonForNonDelivery1);
            this.Controls.Add(this.GrpInspectionDate);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.CmbLogType);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.LsvLogContent);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.LblVersion);
            this.Controls.Add(this.BtnClose);
            this.Controls.Add(this.LblTitle);
            this.Controls.Add(this.LsvLogList);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "LogListForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ログ参照";
            this.Load += new System.EventHandler(this.LogListForm_Load);
            this.GrpInspectionDate.ResumeLayout(false);
            this.GrpInspectionDate.PerformLayout();
            this.GrpReasonForNonDelivery1.ResumeLayout(false);
            this.GrpReasonForNonDelivery1.PerformLayout();
            this.GrpReasonForNonDelivery2.ResumeLayout(false);
            this.GrpReasonForNonDelivery2.PerformLayout();
            this.GrpSortBy.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.PicWaitList)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PicWaitContent)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Label LblVersion;
        internal System.Windows.Forms.Button BtnClose;
        internal System.Windows.Forms.Label LblTitle;
        internal System.Windows.Forms.ComboBox CmbLogType;
        internal System.Windows.Forms.Label label3;
        private System.Windows.Forms.ListView LsvLogContent;
        internal System.Windows.Forms.Label label9;
        internal System.Windows.Forms.Label label1;
        internal System.Windows.Forms.ComboBox CmbReasonForNonDelivery1;
        private System.Windows.Forms.DateTimePicker dtTimePickerFrom;
        private System.Windows.Forms.CheckBox ChkInspectionDate;
        private System.Windows.Forms.GroupBox GrpInspectionDate;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker dtTimePickerTo;
        private System.Windows.Forms.GroupBox GrpReasonForNonDelivery1;
        private System.Windows.Forms.CheckBox ChkReasonForNonDelivery1;
        private System.Windows.Forms.Label LblLogFileCount;
        private System.Windows.Forms.Label LblContentCount;
        internal System.Windows.Forms.Button BtnUpdate;
        internal System.Windows.Forms.Label LblSelectedFile;
        internal System.Windows.Forms.Button BtnJobSelect;
        internal System.Windows.Forms.Button BtnJobClear;
        private System.Windows.Forms.GroupBox GrpReasonForNonDelivery2;
        private System.Windows.Forms.CheckBox ChkReasonForNonDelivery2;
        internal System.Windows.Forms.ComboBox CmbReasonForNonDelivery2;
        private System.Windows.Forms.GroupBox GrpSortBy;
        internal System.Windows.Forms.ComboBox CmbSortBy;
        private System.Windows.Forms.PictureBox PicWaitList;
        private System.Windows.Forms.PictureBox PicWaitContent;
        private System.Windows.Forms.ListView LsvLogList;
    }
}