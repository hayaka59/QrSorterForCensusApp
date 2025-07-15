namespace QrSorterSimulatorApp
{
    partial class MaintenanceForm
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
            this.GroupBox7 = new System.Windows.Forms.GroupBox();
            this.CmbComParityVar = new System.Windows.Forms.ComboBox();
            this.Label27 = new System.Windows.Forms.Label();
            this.CmbComIsParty = new System.Windows.Forms.ComboBox();
            this.Label26 = new System.Windows.Forms.Label();
            this.CmbComStopBit = new System.Windows.Forms.ComboBox();
            this.Label28 = new System.Windows.Forms.Label();
            this.CmbComDataLength = new System.Windows.Forms.ComboBox();
            this.Label24 = new System.Windows.Forms.Label();
            this.CmbComSpeed = new System.Windows.Forms.ComboBox();
            this.Label6 = new System.Windows.Forms.Label();
            this.CmbComPort = new System.Windows.Forms.ComboBox();
            this.Label23 = new System.Windows.Forms.Label();
            this.LblVersion = new System.Windows.Forms.Label();
            this.BtnApply = new System.Windows.Forms.Button();
            this.BtnClose = new System.Windows.Forms.Button();
            this.GroupBox7.SuspendLayout();
            this.SuspendLayout();
            // 
            // GroupBox7
            // 
            this.GroupBox7.Controls.Add(this.CmbComParityVar);
            this.GroupBox7.Controls.Add(this.Label27);
            this.GroupBox7.Controls.Add(this.CmbComIsParty);
            this.GroupBox7.Controls.Add(this.Label26);
            this.GroupBox7.Controls.Add(this.CmbComStopBit);
            this.GroupBox7.Controls.Add(this.Label28);
            this.GroupBox7.Controls.Add(this.CmbComDataLength);
            this.GroupBox7.Controls.Add(this.Label24);
            this.GroupBox7.Controls.Add(this.CmbComSpeed);
            this.GroupBox7.Controls.Add(this.Label6);
            this.GroupBox7.Controls.Add(this.CmbComPort);
            this.GroupBox7.Controls.Add(this.Label23);
            this.GroupBox7.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.GroupBox7.Location = new System.Drawing.Point(304, 155);
            this.GroupBox7.Name = "GroupBox7";
            this.GroupBox7.Size = new System.Drawing.Size(330, 261);
            this.GroupBox7.TabIndex = 247;
            this.GroupBox7.TabStop = false;
            this.GroupBox7.Text = "RS-232C設定";
            // 
            // CmbComParityVar
            // 
            this.CmbComParityVar.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbComParityVar.FormattingEnabled = true;
            this.CmbComParityVar.Location = new System.Drawing.Point(187, 172);
            this.CmbComParityVar.Name = "CmbComParityVar";
            this.CmbComParityVar.Size = new System.Drawing.Size(121, 32);
            this.CmbComParityVar.TabIndex = 97;
            // 
            // Label27
            // 
            this.Label27.BackColor = System.Drawing.Color.Blue;
            this.Label27.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label27.ForeColor = System.Drawing.Color.White;
            this.Label27.Location = new System.Drawing.Point(23, 174);
            this.Label27.Name = "Label27";
            this.Label27.Size = new System.Drawing.Size(163, 30);
            this.Label27.TabIndex = 96;
            this.Label27.Text = "パリティ種別";
            this.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CmbComIsParty
            // 
            this.CmbComIsParty.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbComIsParty.FormattingEnabled = true;
            this.CmbComIsParty.Location = new System.Drawing.Point(187, 137);
            this.CmbComIsParty.Name = "CmbComIsParty";
            this.CmbComIsParty.Size = new System.Drawing.Size(121, 32);
            this.CmbComIsParty.TabIndex = 95;
            // 
            // Label26
            // 
            this.Label26.BackColor = System.Drawing.Color.Blue;
            this.Label26.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label26.ForeColor = System.Drawing.Color.White;
            this.Label26.Location = new System.Drawing.Point(23, 138);
            this.Label26.Name = "Label26";
            this.Label26.Size = new System.Drawing.Size(163, 30);
            this.Label26.TabIndex = 94;
            this.Label26.Text = "パリティ無効／有効";
            this.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CmbComStopBit
            // 
            this.CmbComStopBit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbComStopBit.FormattingEnabled = true;
            this.CmbComStopBit.Location = new System.Drawing.Point(186, 210);
            this.CmbComStopBit.Name = "CmbComStopBit";
            this.CmbComStopBit.Size = new System.Drawing.Size(121, 32);
            this.CmbComStopBit.TabIndex = 93;
            // 
            // Label28
            // 
            this.Label28.BackColor = System.Drawing.Color.Blue;
            this.Label28.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label28.ForeColor = System.Drawing.Color.White;
            this.Label28.Location = new System.Drawing.Point(22, 210);
            this.Label28.Name = "Label28";
            this.Label28.Size = new System.Drawing.Size(163, 30);
            this.Label28.TabIndex = 92;
            this.Label28.Text = "ストップビット";
            this.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CmbComDataLength
            // 
            this.CmbComDataLength.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbComDataLength.FormattingEnabled = true;
            this.CmbComDataLength.Location = new System.Drawing.Point(186, 102);
            this.CmbComDataLength.Name = "CmbComDataLength";
            this.CmbComDataLength.Size = new System.Drawing.Size(121, 32);
            this.CmbComDataLength.TabIndex = 87;
            // 
            // Label24
            // 
            this.Label24.BackColor = System.Drawing.Color.Blue;
            this.Label24.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label24.ForeColor = System.Drawing.Color.White;
            this.Label24.Location = new System.Drawing.Point(22, 103);
            this.Label24.Name = "Label24";
            this.Label24.Size = new System.Drawing.Size(163, 30);
            this.Label24.TabIndex = 86;
            this.Label24.Text = "データ長";
            this.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CmbComSpeed
            // 
            this.CmbComSpeed.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbComSpeed.FormattingEnabled = true;
            this.CmbComSpeed.Location = new System.Drawing.Point(186, 67);
            this.CmbComSpeed.Name = "CmbComSpeed";
            this.CmbComSpeed.Size = new System.Drawing.Size(121, 32);
            this.CmbComSpeed.TabIndex = 85;
            // 
            // Label6
            // 
            this.Label6.BackColor = System.Drawing.Color.Blue;
            this.Label6.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label6.ForeColor = System.Drawing.Color.White;
            this.Label6.Location = new System.Drawing.Point(22, 68);
            this.Label6.Name = "Label6";
            this.Label6.Size = new System.Drawing.Size(163, 30);
            this.Label6.TabIndex = 84;
            this.Label6.Text = "通信速度";
            this.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CmbComPort
            // 
            this.CmbComPort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbComPort.FormattingEnabled = true;
            this.CmbComPort.Location = new System.Drawing.Point(186, 31);
            this.CmbComPort.Name = "CmbComPort";
            this.CmbComPort.Size = new System.Drawing.Size(121, 32);
            this.CmbComPort.TabIndex = 83;
            // 
            // Label23
            // 
            this.Label23.BackColor = System.Drawing.Color.Blue;
            this.Label23.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Label23.ForeColor = System.Drawing.Color.White;
            this.Label23.Location = new System.Drawing.Point(22, 33);
            this.Label23.Name = "Label23";
            this.Label23.Size = new System.Drawing.Size(163, 30);
            this.Label23.TabIndex = 82;
            this.Label23.Text = "ポート";
            this.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // LblVersion
            // 
            this.LblVersion.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblVersion.ForeColor = System.Drawing.Color.Blue;
            this.LblVersion.Location = new System.Drawing.Point(842, 718);
            this.LblVersion.Name = "LblVersion";
            this.LblVersion.Size = new System.Drawing.Size(130, 25);
            this.LblVersion.TabIndex = 249;
            this.LblVersion.Text = "LblVersion";
            this.LblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // BtnApply
            // 
            this.BtnApply.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnApply.Image = global::QrSorterSimulatorApp.Properties.Resources.check;
            this.BtnApply.Location = new System.Drawing.Point(304, 470);
            this.BtnApply.Name = "BtnApply";
            this.BtnApply.Size = new System.Drawing.Size(124, 50);
            this.BtnApply.TabIndex = 250;
            this.BtnApply.Text = "適用";
            this.BtnApply.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnApply.UseVisualStyleBackColor = true;
            this.BtnApply.Click += new System.EventHandler(this.BtnApply_Click);
            // 
            // BtnClose
            // 
            this.BtnClose.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnClose.Image = global::QrSorterSimulatorApp.Properties.Resources.back_arrow;
            this.BtnClose.Location = new System.Drawing.Point(510, 470);
            this.BtnClose.Name = "BtnClose";
            this.BtnClose.Size = new System.Drawing.Size(124, 50);
            this.BtnClose.TabIndex = 248;
            this.BtnClose.Text = "戻る";
            this.BtnClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnClose.UseVisualStyleBackColor = true;
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // MaintenanceForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 761);
            this.ControlBox = false;
            this.Controls.Add(this.BtnApply);
            this.Controls.Add(this.LblVersion);
            this.Controls.Add(this.BtnClose);
            this.Controls.Add(this.GroupBox7);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MaintenanceForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "保守画面★シミュレーター";
            this.Load += new System.EventHandler(this.MaintenanceForm_Load);
            this.GroupBox7.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.GroupBox GroupBox7;
        internal System.Windows.Forms.ComboBox CmbComParityVar;
        internal System.Windows.Forms.Label Label27;
        internal System.Windows.Forms.ComboBox CmbComIsParty;
        internal System.Windows.Forms.Label Label26;
        internal System.Windows.Forms.ComboBox CmbComStopBit;
        internal System.Windows.Forms.Label Label28;
        internal System.Windows.Forms.ComboBox CmbComDataLength;
        internal System.Windows.Forms.Label Label24;
        internal System.Windows.Forms.ComboBox CmbComSpeed;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.ComboBox CmbComPort;
        internal System.Windows.Forms.Label Label23;
        internal System.Windows.Forms.Label LblVersion;
        internal System.Windows.Forms.Button BtnClose;
        internal System.Windows.Forms.Button BtnApply;
    }
}