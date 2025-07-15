namespace QrSorterSimulatorApp
{
    partial class MainForm
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.LblVersion = new System.Windows.Forms.Label();
            this.LsbRecvBox = new System.Windows.Forms.ListBox();
            this.LsbSendBox = new System.Windows.Forms.ListBox();
            this.SerialPortQr = new System.IO.Ports.SerialPort(this.components);
            this.LblError = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.TxtPropertyId = new System.Windows.Forms.TextBox();
            this.dtTimPickPostalDate = new System.Windows.Forms.DateTimePicker();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.TxtUniqueKey = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.CmbJudge = new System.Windows.Forms.ComboBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.CmbErrorCode = new System.Windows.Forms.ComboBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.CmbTray = new System.Windows.Forms.ComboBox();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.TimSendData = new System.Windows.Forms.Timer(this.components);
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.CmbNonDeliveryReasonSorting = new System.Windows.Forms.ComboBox();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.BtnICommand = new System.Windows.Forms.Button();
            this.TxtICommand = new System.Windows.Forms.TextBox();
            this.groupBox11 = new System.Windows.Forms.GroupBox();
            this.BtnJCommand = new System.Windows.Forms.Button();
            this.TxtJCommand = new System.Windows.Forms.TextBox();
            this.BtnDipSw = new System.Windows.Forms.Button();
            this.groupBox12 = new System.Windows.Forms.GroupBox();
            this.CmbTestCount = new System.Windows.Forms.ComboBox();
            this.BtnReset = new System.Windows.Forms.Button();
            this.BtnAutoTest = new System.Windows.Forms.Button();
            this.BtnQrDataSend = new System.Windows.Forms.Button();
            this.BtnInputOutput = new System.Windows.Forms.Button();
            this.BtnConfirmation = new System.Windows.Forms.Button();
            this.BtnError = new System.Windows.Forms.Button();
            this.BtnStop = new System.Windows.Forms.Button();
            this.BtnStart = new System.Windows.Forms.Button();
            this.BtnAutoSend = new System.Windows.Forms.Button();
            this.BtnMaintenance = new System.Windows.Forms.Button();
            this.BtnSendTestData = new System.Windows.Forms.Button();
            this.BtnEnd = new System.Windows.Forms.Button();
            this.groupBox13 = new System.Windows.Forms.GroupBox();
            this.CmbQRDigit = new System.Windows.Forms.ComboBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.groupBox10.SuspendLayout();
            this.groupBox11.SuspendLayout();
            this.groupBox12.SuspendLayout();
            this.groupBox13.SuspendLayout();
            this.SuspendLayout();
            // 
            // LblVersion
            // 
            this.LblVersion.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblVersion.ForeColor = System.Drawing.Color.Blue;
            this.LblVersion.Location = new System.Drawing.Point(822, 731);
            this.LblVersion.Name = "LblVersion";
            this.LblVersion.Size = new System.Drawing.Size(148, 25);
            this.LblVersion.TabIndex = 11;
            this.LblVersion.Text = "LblVersion";
            this.LblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // LsbRecvBox
            // 
            this.LsbRecvBox.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LsbRecvBox.FormattingEnabled = true;
            this.LsbRecvBox.ItemHeight = 24;
            this.LsbRecvBox.Location = new System.Drawing.Point(12, 22);
            this.LsbRecvBox.Name = "LsbRecvBox";
            this.LsbRecvBox.Size = new System.Drawing.Size(549, 196);
            this.LsbRecvBox.TabIndex = 12;
            // 
            // LsbSendBox
            // 
            this.LsbSendBox.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LsbSendBox.FormattingEnabled = true;
            this.LsbSendBox.ItemHeight = 24;
            this.LsbSendBox.Location = new System.Drawing.Point(11, 22);
            this.LsbSendBox.Name = "LsbSendBox";
            this.LsbSendBox.Size = new System.Drawing.Size(549, 196);
            this.LsbSendBox.TabIndex = 13;
            // 
            // LblError
            // 
            this.LblError.BackColor = System.Drawing.Color.LightCoral;
            this.LblError.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblError.ForeColor = System.Drawing.Color.Blue;
            this.LblError.Location = new System.Drawing.Point(12, 701);
            this.LblError.Name = "LblError";
            this.LblError.Size = new System.Drawing.Size(776, 45);
            this.LblError.TabIndex = 318;
            this.LblError.Text = "LblError";
            this.LblError.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.LblError.Visible = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.TxtPropertyId);
            this.groupBox1.Location = new System.Drawing.Point(32, 94);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(97, 60);
            this.groupBox1.TabIndex = 321;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "物件ID（5桁）";
            // 
            // TxtPropertyId
            // 
            this.TxtPropertyId.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TxtPropertyId.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.TxtPropertyId.Location = new System.Drawing.Point(6, 18);
            this.TxtPropertyId.MaxLength = 5;
            this.TxtPropertyId.Name = "TxtPropertyId";
            this.TxtPropertyId.Size = new System.Drawing.Size(81, 31);
            this.TxtPropertyId.TabIndex = 0;
            this.TxtPropertyId.Text = "A8657";
            this.TxtPropertyId.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // dtTimPickPostalDate
            // 
            this.dtTimPickPostalDate.CalendarFont = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.dtTimPickPostalDate.Location = new System.Drawing.Point(11, 23);
            this.dtTimPickPostalDate.Name = "dtTimPickPostalDate";
            this.dtTimPickPostalDate.Size = new System.Drawing.Size(126, 19);
            this.dtTimPickPostalDate.TabIndex = 322;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dtTimPickPostalDate);
            this.groupBox2.Location = new System.Drawing.Point(145, 94);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(150, 60);
            this.groupBox2.TabIndex = 322;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "局出し日（8桁）";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.TxtUniqueKey);
            this.groupBox3.Location = new System.Drawing.Point(32, 160);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(263, 60);
            this.groupBox3.TabIndex = 323;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "ユニークキー（17桁）の内、9桁を入力";
            // 
            // TxtUniqueKey
            // 
            this.TxtUniqueKey.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TxtUniqueKey.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.TxtUniqueKey.Location = new System.Drawing.Point(6, 18);
            this.TxtUniqueKey.MaxLength = 17;
            this.TxtUniqueKey.Name = "TxtUniqueKey";
            this.TxtUniqueKey.Size = new System.Drawing.Size(244, 31);
            this.TxtUniqueKey.TabIndex = 0;
            this.TxtUniqueKey.Text = "000000001";
            this.TxtUniqueKey.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.CmbJudge);
            this.groupBox4.Location = new System.Drawing.Point(32, 226);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(73, 60);
            this.groupBox4.TabIndex = 324;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "判定";
            // 
            // CmbJudge
            // 
            this.CmbJudge.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbJudge.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CmbJudge.FormattingEnabled = true;
            this.CmbJudge.Location = new System.Drawing.Point(6, 18);
            this.CmbJudge.Name = "CmbJudge";
            this.CmbJudge.Size = new System.Drawing.Size(56, 32);
            this.CmbJudge.TabIndex = 0;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.CmbErrorCode);
            this.groupBox5.Location = new System.Drawing.Point(22, 518);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(83, 60);
            this.groupBox5.TabIndex = 325;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "エラーコード";
            // 
            // CmbErrorCode
            // 
            this.CmbErrorCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbErrorCode.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CmbErrorCode.FormattingEnabled = true;
            this.CmbErrorCode.Location = new System.Drawing.Point(6, 18);
            this.CmbErrorCode.Name = "CmbErrorCode";
            this.CmbErrorCode.Size = new System.Drawing.Size(63, 32);
            this.CmbErrorCode.TabIndex = 1;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.CmbTray);
            this.groupBox6.Location = new System.Drawing.Point(111, 226);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(74, 60);
            this.groupBox6.TabIndex = 326;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "トレイ情報";
            // 
            // CmbTray
            // 
            this.CmbTray.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbTray.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CmbTray.FormattingEnabled = true;
            this.CmbTray.Location = new System.Drawing.Point(6, 18);
            this.CmbTray.Name = "CmbTray";
            this.CmbTray.Size = new System.Drawing.Size(56, 32);
            this.CmbTray.TabIndex = 1;
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.LsbSendBox);
            this.groupBox7.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.groupBox7.Location = new System.Drawing.Point(382, 21);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(570, 230);
            this.groupBox7.TabIndex = 327;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "送信データ";
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.LsbRecvBox);
            this.groupBox8.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.groupBox8.Location = new System.Drawing.Point(382, 259);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(570, 230);
            this.groupBox8.TabIndex = 328;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "受信データ";
            // 
            // TimSendData
            // 
            this.TimSendData.Tick += new System.EventHandler(this.TimSendData_Tick);
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.CmbNonDeliveryReasonSorting);
            this.groupBox9.Location = new System.Drawing.Point(713, 726);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(87, 60);
            this.groupBox9.TabIndex = 327;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "不着事由";
            this.groupBox9.Visible = false;
            // 
            // CmbNonDeliveryReasonSorting
            // 
            this.CmbNonDeliveryReasonSorting.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbNonDeliveryReasonSorting.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CmbNonDeliveryReasonSorting.FormattingEnabled = true;
            this.CmbNonDeliveryReasonSorting.Location = new System.Drawing.Point(6, 18);
            this.CmbNonDeliveryReasonSorting.Name = "CmbNonDeliveryReasonSorting";
            this.CmbNonDeliveryReasonSorting.Size = new System.Drawing.Size(72, 32);
            this.CmbNonDeliveryReasonSorting.TabIndex = 1;
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.BtnICommand);
            this.groupBox10.Controls.Add(this.TxtICommand);
            this.groupBox10.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.groupBox10.Location = new System.Drawing.Point(534, 492);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(196, 60);
            this.groupBox10.TabIndex = 335;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "アワーメーター";
            // 
            // BtnICommand
            // 
            this.BtnICommand.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnICommand.Location = new System.Drawing.Point(129, 14);
            this.BtnICommand.Name = "BtnICommand";
            this.BtnICommand.Size = new System.Drawing.Size(57, 40);
            this.BtnICommand.TabIndex = 337;
            this.BtnICommand.Text = "I";
            this.BtnICommand.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnICommand.UseVisualStyleBackColor = true;
            this.BtnICommand.Click += new System.EventHandler(this.BtnICommand_Click);
            // 
            // TxtICommand
            // 
            this.TxtICommand.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TxtICommand.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.TxtICommand.Location = new System.Drawing.Point(6, 18);
            this.TxtICommand.MaxLength = 17;
            this.TxtICommand.Name = "TxtICommand";
            this.TxtICommand.Size = new System.Drawing.Size(117, 31);
            this.TxtICommand.TabIndex = 0;
            this.TxtICommand.Text = "12345";
            this.TxtICommand.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // groupBox11
            // 
            this.groupBox11.Controls.Add(this.BtnJCommand);
            this.groupBox11.Controls.Add(this.TxtJCommand);
            this.groupBox11.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.groupBox11.Location = new System.Drawing.Point(756, 492);
            this.groupBox11.Name = "groupBox11";
            this.groupBox11.Size = new System.Drawing.Size(196, 60);
            this.groupBox11.TabIndex = 336;
            this.groupBox11.TabStop = false;
            this.groupBox11.Text = "トータルカウンタ";
            // 
            // BtnJCommand
            // 
            this.BtnJCommand.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnJCommand.Location = new System.Drawing.Point(129, 14);
            this.BtnJCommand.Name = "BtnJCommand";
            this.BtnJCommand.Size = new System.Drawing.Size(57, 40);
            this.BtnJCommand.TabIndex = 338;
            this.BtnJCommand.Text = "J";
            this.BtnJCommand.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnJCommand.UseVisualStyleBackColor = true;
            this.BtnJCommand.Click += new System.EventHandler(this.BtnJCommand_Click);
            // 
            // TxtJCommand
            // 
            this.TxtJCommand.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TxtJCommand.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.TxtJCommand.Location = new System.Drawing.Point(6, 18);
            this.TxtJCommand.MaxLength = 17;
            this.TxtJCommand.Name = "TxtJCommand";
            this.TxtJCommand.Size = new System.Drawing.Size(117, 31);
            this.TxtJCommand.TabIndex = 0;
            this.TxtJCommand.Text = "123456789";
            this.TxtJCommand.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // BtnDipSw
            // 
            this.BtnDipSw.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnDipSw.Location = new System.Drawing.Point(38, 592);
            this.BtnDipSw.Name = "BtnDipSw";
            this.BtnDipSw.Size = new System.Drawing.Size(257, 45);
            this.BtnDipSw.TabIndex = 337;
            this.BtnDipSw.Text = "DIP-SW（T）";
            this.BtnDipSw.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnDipSw.UseVisualStyleBackColor = true;
            this.BtnDipSw.Click += new System.EventHandler(this.BtnDipSw_Click);
            // 
            // groupBox12
            // 
            this.groupBox12.Controls.Add(this.CmbTestCount);
            this.groupBox12.Location = new System.Drawing.Point(686, 634);
            this.groupBox12.Name = "groupBox12";
            this.groupBox12.Size = new System.Drawing.Size(83, 60);
            this.groupBox12.TabIndex = 340;
            this.groupBox12.TabStop = false;
            this.groupBox12.Text = "テスト通数";
            // 
            // CmbTestCount
            // 
            this.CmbTestCount.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbTestCount.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CmbTestCount.FormattingEnabled = true;
            this.CmbTestCount.Location = new System.Drawing.Point(6, 18);
            this.CmbTestCount.Name = "CmbTestCount";
            this.CmbTestCount.Size = new System.Drawing.Size(63, 32);
            this.CmbTestCount.TabIndex = 1;
            // 
            // BtnReset
            // 
            this.BtnReset.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnReset.Image = global::QrSorterSimulatorApp.Properties.Resources.peke_icon;
            this.BtnReset.Location = new System.Drawing.Point(38, 453);
            this.BtnReset.Name = "BtnReset";
            this.BtnReset.Size = new System.Drawing.Size(257, 45);
            this.BtnReset.TabIndex = 341;
            this.BtnReset.Text = "リセット（F）";
            this.BtnReset.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnReset.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnReset.UseVisualStyleBackColor = true;
            this.BtnReset.Click += new System.EventHandler(this.BtnReset_Click);
            // 
            // BtnAutoTest
            // 
            this.BtnAutoTest.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnAutoTest.Image = global::QrSorterSimulatorApp.Properties.Resources.detective;
            this.BtnAutoTest.Location = new System.Drawing.Point(686, 584);
            this.BtnAutoTest.Name = "BtnAutoTest";
            this.BtnAutoTest.Size = new System.Drawing.Size(257, 45);
            this.BtnAutoTest.TabIndex = 339;
            this.BtnAutoTest.Text = "自動テスト";
            this.BtnAutoTest.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnAutoTest.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnAutoTest.UseVisualStyleBackColor = true;
            this.BtnAutoTest.Click += new System.EventHandler(this.BtnAutoTest_Click);
            // 
            // BtnQrDataSend
            // 
            this.BtnQrDataSend.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnQrDataSend.Image = global::QrSorterSimulatorApp.Properties.Resources.qr_code;
            this.BtnQrDataSend.Location = new System.Drawing.Point(393, 584);
            this.BtnQrDataSend.Name = "BtnQrDataSend";
            this.BtnQrDataSend.Size = new System.Drawing.Size(257, 45);
            this.BtnQrDataSend.TabIndex = 338;
            this.BtnQrDataSend.Text = "QRデータ（L）";
            this.BtnQrDataSend.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnQrDataSend.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnQrDataSend.UseVisualStyleBackColor = true;
            this.BtnQrDataSend.Click += new System.EventHandler(this.BtnQrDataSend_Click);
            // 
            // BtnInputOutput
            // 
            this.BtnInputOutput.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnInputOutput.Image = global::QrSorterSimulatorApp.Properties.Resources.bubble;
            this.BtnInputOutput.Location = new System.Drawing.Point(38, 650);
            this.BtnInputOutput.Name = "BtnInputOutput";
            this.BtnInputOutput.Size = new System.Drawing.Size(257, 45);
            this.BtnInputOutput.TabIndex = 334;
            this.BtnInputOutput.Text = "I/O（K）";
            this.BtnInputOutput.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnInputOutput.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnInputOutput.UseVisualStyleBackColor = true;
            this.BtnInputOutput.Click += new System.EventHandler(this.BtnInputOutput_Click);
            // 
            // BtnConfirmation
            // 
            this.BtnConfirmation.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnConfirmation.Image = global::QrSorterSimulatorApp.Properties.Resources.bubble;
            this.BtnConfirmation.Location = new System.Drawing.Point(38, 299);
            this.BtnConfirmation.Name = "BtnConfirmation";
            this.BtnConfirmation.Size = new System.Drawing.Size(257, 45);
            this.BtnConfirmation.TabIndex = 333;
            this.BtnConfirmation.Text = "確認（A）";
            this.BtnConfirmation.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnConfirmation.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnConfirmation.UseVisualStyleBackColor = true;
            this.BtnConfirmation.Click += new System.EventHandler(this.BtnConfirmation_Click);
            // 
            // BtnError
            // 
            this.BtnError.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnError.Image = global::QrSorterSimulatorApp.Properties.Resources.warning;
            this.BtnError.Location = new System.Drawing.Point(111, 526);
            this.BtnError.Name = "BtnError";
            this.BtnError.Size = new System.Drawing.Size(246, 45);
            this.BtnError.TabIndex = 332;
            this.BtnError.Text = "エラー送信（E）";
            this.BtnError.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnError.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnError.UseVisualStyleBackColor = true;
            this.BtnError.Click += new System.EventHandler(this.BtnError_Click);
            // 
            // BtnStop
            // 
            this.BtnStop.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnStop.Image = global::QrSorterSimulatorApp.Properties.Resources.standing;
            this.BtnStop.Location = new System.Drawing.Point(38, 402);
            this.BtnStop.Name = "BtnStop";
            this.BtnStop.Size = new System.Drawing.Size(257, 45);
            this.BtnStop.TabIndex = 331;
            this.BtnStop.Text = "停止（C）";
            this.BtnStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnStop.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnStop.UseVisualStyleBackColor = true;
            this.BtnStop.Click += new System.EventHandler(this.BtnStop_Click);
            // 
            // BtnStart
            // 
            this.BtnStart.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnStart.Image = global::QrSorterSimulatorApp.Properties.Resources.running_icon;
            this.BtnStart.Location = new System.Drawing.Point(38, 351);
            this.BtnStart.Name = "BtnStart";
            this.BtnStart.Size = new System.Drawing.Size(257, 45);
            this.BtnStart.TabIndex = 330;
            this.BtnStart.Text = "開始（B）";
            this.BtnStart.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnStart.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnStart.UseVisualStyleBackColor = true;
            this.BtnStart.Click += new System.EventHandler(this.BtnStart_Click);
            // 
            // BtnAutoSend
            // 
            this.BtnAutoSend.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnAutoSend.Image = global::QrSorterSimulatorApp.Properties.Resources.repeat;
            this.BtnAutoSend.Location = new System.Drawing.Point(393, 642);
            this.BtnAutoSend.Name = "BtnAutoSend";
            this.BtnAutoSend.Size = new System.Drawing.Size(257, 45);
            this.BtnAutoSend.TabIndex = 329;
            this.BtnAutoSend.Text = "連続送信";
            this.BtnAutoSend.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnAutoSend.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnAutoSend.UseVisualStyleBackColor = true;
            this.BtnAutoSend.Click += new System.EventHandler(this.BtnAutoSend_Click);
            // 
            // BtnMaintenance
            // 
            this.BtnMaintenance.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnMaintenance.Image = global::QrSorterSimulatorApp.Properties.Resources.settei;
            this.BtnMaintenance.Location = new System.Drawing.Point(393, 501);
            this.BtnMaintenance.Name = "BtnMaintenance";
            this.BtnMaintenance.Size = new System.Drawing.Size(113, 45);
            this.BtnMaintenance.TabIndex = 320;
            this.BtnMaintenance.Text = "設定";
            this.BtnMaintenance.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnMaintenance.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnMaintenance.UseVisualStyleBackColor = true;
            this.BtnMaintenance.Click += new System.EventHandler(this.BtnMaintenance_Click);
            // 
            // BtnSendTestData
            // 
            this.BtnSendTestData.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnSendTestData.Image = global::QrSorterSimulatorApp.Properties.Resources.arrow;
            this.BtnSendTestData.Location = new System.Drawing.Point(34, 32);
            this.BtnSendTestData.Name = "BtnSendTestData";
            this.BtnSendTestData.Size = new System.Drawing.Size(257, 45);
            this.BtnSendTestData.TabIndex = 319;
            this.BtnSendTestData.Text = "テストデータ送信";
            this.BtnSendTestData.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnSendTestData.UseVisualStyleBackColor = true;
            this.BtnSendTestData.Click += new System.EventHandler(this.BtnSendTestData_Click);
            // 
            // BtnEnd
            // 
            this.BtnEnd.Font = new System.Drawing.Font("メイリオ", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnEnd.Image = global::QrSorterSimulatorApp.Properties.Resources.exit_icon_small;
            this.BtnEnd.Location = new System.Drawing.Point(821, 682);
            this.BtnEnd.Name = "BtnEnd";
            this.BtnEnd.Size = new System.Drawing.Size(149, 45);
            this.BtnEnd.TabIndex = 2;
            this.BtnEnd.Text = " 終了";
            this.BtnEnd.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnEnd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnEnd.UseVisualStyleBackColor = true;
            this.BtnEnd.Click += new System.EventHandler(this.BtnEnd_Click);
            // 
            // groupBox13
            // 
            this.groupBox13.Controls.Add(this.CmbQRDigit);
            this.groupBox13.Location = new System.Drawing.Point(221, 226);
            this.groupBox13.Name = "groupBox13";
            this.groupBox13.Size = new System.Drawing.Size(74, 60);
            this.groupBox13.TabIndex = 327;
            this.groupBox13.TabStop = false;
            this.groupBox13.Text = "QR桁数";
            // 
            // CmbQRDigit
            // 
            this.CmbQRDigit.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.CmbQRDigit.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CmbQRDigit.FormattingEnabled = true;
            this.CmbQRDigit.Location = new System.Drawing.Point(6, 18);
            this.CmbQRDigit.Name = "CmbQRDigit";
            this.CmbQRDigit.Size = new System.Drawing.Size(56, 32);
            this.CmbQRDigit.TabIndex = 1;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 761);
            this.ControlBox = false;
            this.Controls.Add(this.groupBox13);
            this.Controls.Add(this.BtnReset);
            this.Controls.Add(this.groupBox12);
            this.Controls.Add(this.BtnAutoTest);
            this.Controls.Add(this.BtnQrDataSend);
            this.Controls.Add(this.BtnDipSw);
            this.Controls.Add(this.groupBox11);
            this.Controls.Add(this.groupBox10);
            this.Controls.Add(this.BtnInputOutput);
            this.Controls.Add(this.BtnConfirmation);
            this.Controls.Add(this.BtnError);
            this.Controls.Add(this.BtnStop);
            this.Controls.Add(this.BtnStart);
            this.Controls.Add(this.groupBox9);
            this.Controls.Add(this.BtnAutoSend);
            this.Controls.Add(this.groupBox8);
            this.Controls.Add(this.groupBox7);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.BtnMaintenance);
            this.Controls.Add(this.BtnSendTestData);
            this.Controls.Add(this.LblError);
            this.Controls.Add(this.LblVersion);
            this.Controls.Add(this.BtnEnd);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "■メインメニュー画面■";
            this.Activated += new System.EventHandler(this.MainForm_Activated);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            this.groupBox8.ResumeLayout(false);
            this.groupBox9.ResumeLayout(false);
            this.groupBox10.ResumeLayout(false);
            this.groupBox10.PerformLayout();
            this.groupBox11.ResumeLayout(false);
            this.groupBox11.PerformLayout();
            this.groupBox12.ResumeLayout(false);
            this.groupBox13.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button BtnEnd;
        internal System.Windows.Forms.Label LblVersion;
        private System.Windows.Forms.ListBox LsbRecvBox;
        private System.Windows.Forms.ListBox LsbSendBox;
        internal System.IO.Ports.SerialPort SerialPortQr;
        internal System.Windows.Forms.Label LblError;
        private System.Windows.Forms.Button BtnSendTestData;
        private System.Windows.Forms.Button BtnMaintenance;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox TxtPropertyId;
        private System.Windows.Forms.DateTimePicker dtTimPickPostalDate;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox TxtUniqueKey;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.ComboBox CmbJudge;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.ComboBox CmbErrorCode;
        private System.Windows.Forms.ComboBox CmbTray;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.GroupBox groupBox7;
        private System.Windows.Forms.GroupBox groupBox8;
        private System.Windows.Forms.Timer TimSendData;
        private System.Windows.Forms.Button BtnAutoSend;
        private System.Windows.Forms.GroupBox groupBox9;
        private System.Windows.Forms.ComboBox CmbNonDeliveryReasonSorting;
        private System.Windows.Forms.Button BtnStart;
        private System.Windows.Forms.Button BtnStop;
        private System.Windows.Forms.Button BtnError;
        private System.Windows.Forms.Button BtnConfirmation;
        private System.Windows.Forms.Button BtnInputOutput;
        private System.Windows.Forms.GroupBox groupBox10;
        private System.Windows.Forms.TextBox TxtICommand;
        private System.Windows.Forms.GroupBox groupBox11;
        private System.Windows.Forms.TextBox TxtJCommand;
        private System.Windows.Forms.Button BtnICommand;
        private System.Windows.Forms.Button BtnJCommand;
        private System.Windows.Forms.Button BtnDipSw;
        private System.Windows.Forms.Button BtnQrDataSend;
        private System.Windows.Forms.Button BtnAutoTest;
        private System.Windows.Forms.GroupBox groupBox12;
        private System.Windows.Forms.ComboBox CmbTestCount;
        private System.Windows.Forms.Button BtnReset;
        private System.Windows.Forms.GroupBox groupBox13;
        private System.Windows.Forms.ComboBox CmbQRDigit;
    }
}

