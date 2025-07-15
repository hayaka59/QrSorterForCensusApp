namespace QrSorterInspectionApp
{
    partial class MenuForm
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
            this.components = new System.ComponentModel.Container();
            this.LblVersion = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnEnd = new System.Windows.Forms.Button();
            this.BtnLogOut = new System.Windows.Forms.Button();
            this.BtnLogManagement = new System.Windows.Forms.Button();
            this.BtnMaintenance = new System.Windows.Forms.Button();
            this.BtnAccountSet = new System.Windows.Forms.Button();
            this.BtnSetting = new System.Windows.Forms.Button();
            this.BtnQrSorterInspect = new System.Windows.Forms.Button();
            this.LblStatus = new System.Windows.Forms.Label();
            this.LblUserInfo = new System.Windows.Forms.Label();
            this.SerialPort = new System.IO.Ports.SerialPort(this.components);
            this.SuspendLayout();
            // 
            // LblVersion
            // 
            this.LblVersion.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblVersion.ForeColor = System.Drawing.Color.Blue;
            this.LblVersion.Location = new System.Drawing.Point(1762, 996);
            this.LblVersion.Name = "LblVersion";
            this.LblVersion.Size = new System.Drawing.Size(130, 25);
            this.LblVersion.TabIndex = 12;
            this.LblVersion.Text = "LblVersion";
            this.LblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.LblVersion.DoubleClick += new System.EventHandler(this.LblVersion_DoubleClick);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.CornflowerBlue;
            this.label1.Font = new System.Drawing.Font("メイリオ", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(0, 1);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(1903, 45);
            this.label1.TabIndex = 11;
            this.label1.Text = "メニュー画面";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // BtnEnd
            // 
            this.BtnEnd.Font = new System.Drawing.Font("メイリオ", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnEnd.Image = global::QrSorterInspectionApp.Properties.Resources.exit;
            this.BtnEnd.ImageAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnEnd.Location = new System.Drawing.Point(961, 742);
            this.BtnEnd.Name = "BtnEnd";
            this.BtnEnd.Size = new System.Drawing.Size(750, 200);
            this.BtnEnd.TabIndex = 7;
            this.BtnEnd.Text = "終了";
            this.BtnEnd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.BtnEnd.UseVisualStyleBackColor = true;
            this.BtnEnd.Click += new System.EventHandler(this.BtnEnd_Click);
            // 
            // BtnLogOut
            // 
            this.BtnLogOut.Font = new System.Drawing.Font("メイリオ", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnLogOut.Image = global::QrSorterInspectionApp.Properties.Resources.login_big;
            this.BtnLogOut.ImageAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnLogOut.Location = new System.Drawing.Point(194, 742);
            this.BtnLogOut.Name = "BtnLogOut";
            this.BtnLogOut.Size = new System.Drawing.Size(750, 200);
            this.BtnLogOut.TabIndex = 6;
            this.BtnLogOut.Text = "ログアウト";
            this.BtnLogOut.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.BtnLogOut.UseVisualStyleBackColor = true;
            this.BtnLogOut.Click += new System.EventHandler(this.BtnLogOut_Click);
            // 
            // BtnLogManagement
            // 
            this.BtnLogManagement.Font = new System.Drawing.Font("メイリオ", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnLogManagement.Image = global::QrSorterInspectionApp.Properties.Resources.download;
            this.BtnLogManagement.ImageAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnLogManagement.Location = new System.Drawing.Point(961, 520);
            this.BtnLogManagement.Name = "BtnLogManagement";
            this.BtnLogManagement.Size = new System.Drawing.Size(750, 200);
            this.BtnLogManagement.TabIndex = 5;
            this.BtnLogManagement.Text = "ログ管理";
            this.BtnLogManagement.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.BtnLogManagement.UseVisualStyleBackColor = true;
            this.BtnLogManagement.Click += new System.EventHandler(this.BtnLogManagement_Click);
            // 
            // BtnMaintenance
            // 
            this.BtnMaintenance.Font = new System.Drawing.Font("メイリオ", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnMaintenance.Image = global::QrSorterInspectionApp.Properties.Resources.maintenance_icon;
            this.BtnMaintenance.ImageAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnMaintenance.Location = new System.Drawing.Point(194, 520);
            this.BtnMaintenance.Name = "BtnMaintenance";
            this.BtnMaintenance.Size = new System.Drawing.Size(750, 200);
            this.BtnMaintenance.TabIndex = 4;
            this.BtnMaintenance.Text = "保守";
            this.BtnMaintenance.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.BtnMaintenance.UseVisualStyleBackColor = true;
            this.BtnMaintenance.Click += new System.EventHandler(this.BtnMaintenance_Click);
            // 
            // BtnAccountSet
            // 
            this.BtnAccountSet.Font = new System.Drawing.Font("メイリオ", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnAccountSet.Image = global::QrSorterInspectionApp.Properties.Resources.sv_op;
            this.BtnAccountSet.ImageAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnAccountSet.Location = new System.Drawing.Point(961, 300);
            this.BtnAccountSet.Name = "BtnAccountSet";
            this.BtnAccountSet.Size = new System.Drawing.Size(750, 200);
            this.BtnAccountSet.TabIndex = 3;
            this.BtnAccountSet.Text = "SV・OP";
            this.BtnAccountSet.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.BtnAccountSet.UseVisualStyleBackColor = true;
            this.BtnAccountSet.Click += new System.EventHandler(this.BtnAccountSet_Click);
            // 
            // BtnSetting
            // 
            this.BtnSetting.Font = new System.Drawing.Font("メイリオ", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnSetting.Image = global::QrSorterInspectionApp.Properties.Resources.setting;
            this.BtnSetting.ImageAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnSetting.Location = new System.Drawing.Point(194, 300);
            this.BtnSetting.Name = "BtnSetting";
            this.BtnSetting.Size = new System.Drawing.Size(750, 200);
            this.BtnSetting.TabIndex = 2;
            this.BtnSetting.Text = "設定";
            this.BtnSetting.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.BtnSetting.UseVisualStyleBackColor = true;
            this.BtnSetting.Click += new System.EventHandler(this.BtnSetting_Click);
            // 
            // BtnQrSorterInspect
            // 
            this.BtnQrSorterInspect.Font = new System.Drawing.Font("メイリオ", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnQrSorterInspect.Image = global::QrSorterInspectionApp.Properties.Resources.qr_code;
            this.BtnQrSorterInspect.ImageAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnQrSorterInspect.Location = new System.Drawing.Point(194, 83);
            this.BtnQrSorterInspect.Name = "BtnQrSorterInspect";
            this.BtnQrSorterInspect.Size = new System.Drawing.Size(1517, 200);
            this.BtnQrSorterInspect.TabIndex = 1;
            this.BtnQrSorterInspect.Text = "QRソータ検査";
            this.BtnQrSorterInspect.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.BtnQrSorterInspect.UseVisualStyleBackColor = true;
            this.BtnQrSorterInspect.Click += new System.EventHandler(this.BtnQrSorterInspect_Click);
            // 
            // LblStatus
            // 
            this.LblStatus.BackColor = System.Drawing.Color.LightCoral;
            this.LblStatus.Font = new System.Drawing.Font("メイリオ", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblStatus.ForeColor = System.Drawing.Color.Blue;
            this.LblStatus.Location = new System.Drawing.Point(194, 959);
            this.LblStatus.Name = "LblStatus";
            this.LblStatus.Size = new System.Drawing.Size(1517, 62);
            this.LblStatus.TabIndex = 265;
            this.LblStatus.Text = "LblStatus";
            this.LblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.LblStatus.Visible = false;
            // 
            // LblUserInfo
            // 
            this.LblUserInfo.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblUserInfo.ForeColor = System.Drawing.Color.Blue;
            this.LblUserInfo.Location = new System.Drawing.Point(1717, 824);
            this.LblUserInfo.Name = "LblUserInfo";
            this.LblUserInfo.Size = new System.Drawing.Size(175, 118);
            this.LblUserInfo.TabIndex = 266;
            this.LblUserInfo.Text = "LblUserInfo";
            this.LblUserInfo.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            this.LblUserInfo.Visible = false;
            // 
            // MenuForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1904, 1041);
            this.ControlBox = false;
            this.Controls.Add(this.LblUserInfo);
            this.Controls.Add(this.LblStatus);
            this.Controls.Add(this.BtnEnd);
            this.Controls.Add(this.BtnLogOut);
            this.Controls.Add(this.BtnLogManagement);
            this.Controls.Add(this.BtnMaintenance);
            this.Controls.Add(this.BtnAccountSet);
            this.Controls.Add(this.BtnSetting);
            this.Controls.Add(this.BtnQrSorterInspect);
            this.Controls.Add(this.LblVersion);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MenuForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "メニュー画面";
            this.Load += new System.EventHandler(this.MenuForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Label LblVersion;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button BtnQrSorterInspect;
        private System.Windows.Forms.Button BtnSetting;
        private System.Windows.Forms.Button BtnAccountSet;
        private System.Windows.Forms.Button BtnLogManagement;
        private System.Windows.Forms.Button BtnMaintenance;
        private System.Windows.Forms.Button BtnEnd;
        private System.Windows.Forms.Button BtnLogOut;
        internal System.Windows.Forms.Label LblStatus;
        internal System.Windows.Forms.Label LblUserInfo;
        internal System.IO.Ports.SerialPort SerialPort;
    }
}