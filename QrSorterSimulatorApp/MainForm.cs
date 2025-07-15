using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QrSorterSimulatorApp
{
    public partial class MainForm : Form
    {
        private delegate void Delegate_RcvDataToTextBox(string data);

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                // バージョン表示
                LblVersion.Text = PubConstClass.DEF_VERSION;
                PubConstClass.objSyncHist = new object();
                CommonModule.OutPutLogFile("〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓");
                CommonModule.OutPutLogFile("【" + "シミュレータ" + "】を起動しました。");
                CommonModule.OutPutLogFile("■シミュレータ★バージョン「" + PubConstClass.DEF_VERSION + "」");

                CommonModule.ReadSystemDefinition();
                String sTitle = "";
                #region シリアルポートの設定
                // データ受信イベントの設定
                SerialPortQr.DataReceived += new SerialDataReceivedEventHandler(SerialPortQr_DataReceived);
                // シリアルポート名の設定
                SerialPortQr.PortName = PubConstClass.pblComPort;
                sTitle += PubConstClass.pblComPort + "／";
                // シリアルポートの通信速度指定
                switch (PubConstClass.pblComSpeed)
                {
                    case "0":
                        {
                            SerialPortQr.BaudRate = 4800;
                            sTitle += "4800bps／";
                            break;
                        }

                    case "1":
                        {
                            SerialPortQr.BaudRate = 9600;
                            sTitle += "9600bps／";
                            break;
                        }

                    case "2":
                        {
                            SerialPortQr.BaudRate = 19200;
                            sTitle += "19200bps／";
                            break;
                        }

                    case "3":
                        {
                            SerialPortQr.BaudRate = 38400;
                            sTitle += "38400bps／";
                            break;
                        }

                    case "4":
                        {
                            SerialPortQr.BaudRate = 57600;
                            sTitle += "57600bps／";
                            break;
                        }

                    case "5":
                        {
                            SerialPortQr.BaudRate = 115200;
                            sTitle += "115200bps／";
                            break;
                        }

                    default:
                        {
                            SerialPortQr.BaudRate = 38400;
                            break;
                        }
                }
                // シリアルポートのパリティ指定
                switch (PubConstClass.pblComParityVar)
                {
                    case "0":
                        {
                            SerialPortQr.Parity = Parity.Odd;
                            sTitle += "奇数／";
                            break;
                        }

                    case "1":
                        {
                            SerialPortQr.Parity = Parity.Even;
                            sTitle += "偶数／";
                            break;
                        }

                    default:
                        {
                            SerialPortQr.Parity = Parity.Even;
                            break;
                        }
                }
                // シリアルポートのパリティ有無
                if (PubConstClass.pblComIsParity == "0")
                {
                    SerialPortQr.Parity = Parity.None;
                    sTitle += "パリティ無し／";
                }                    
                // シリアルポートのビット数指定
                switch (PubConstClass.pblComDataLength)
                {
                    case "0":
                        {
                            SerialPortQr.DataBits = 8;
                            sTitle += "8ビット／";
                            break;
                        }

                    case "1":
                        {
                            SerialPortQr.DataBits = 7;
                            sTitle += "7ビット／";
                            break;
                        }

                    default:
                        {
                            SerialPortQr.DataBits = 8;
                            break;
                        }
                }
                // シリアルポートのストップビット指定
                switch (PubConstClass.pblComStopBit)
                {
                    case "0":
                        {
                            SerialPortQr.StopBits = StopBits.One;
                            sTitle += "ストップビット１ ］";
                            break;
                        }

                    case "1":
                        {
                            SerialPortQr.StopBits = StopBits.Two;
                            sTitle += "ストップビット２ ］";
                            break;
                        }

                    default:
                        {
                            SerialPortQr.StopBits = StopBits.One;
                            break;
                        }
                }
                #endregion

                PubConstClass.pblMainFormTitle = "【メインメニュー画面】 ［ " + sTitle;
                // 判定
                CmbJudge.Items.Clear();
                CmbJudge.Items.Add("OK");
                CmbJudge.Items.Add("NG");
                CmbJudge.SelectedIndex = 0;
                // エラーコード
                CmbErrorCode.Items.Clear();
                for (int i = 0; i < 300; i++){
                    CmbErrorCode.Items.Add(i.ToString("000"));
                }
                CmbErrorCode.SelectedIndex = 0;
                // トレイ情報
                CmbTray.Items.Clear();
                CmbTray.Items.Add("1");
                CmbTray.Items.Add("2");
                CmbTray.Items.Add("3");
                CmbTray.Items.Add("4");
                CmbTray.Items.Add("5");
                CmbTray.Items.Add("E");
                CmbTray.Items.Add("6");
                CmbTray.Items.Add("7");
                CmbTray.Items.Add("8");
                CmbTray.Items.Add("9");                
                CmbTray.SelectedIndex = 0;

                CmbQRDigit.Items.Clear();
                for (int i = 0; i <= 128 - 31; i++)
                {
                    CmbQRDigit.Items.Add((i + 31).ToString("000"));
                }   
                CmbQRDigit.SelectedIndex = 0;

                #region 不着事由区分
                CommonModule.ReadNonDeliveryList();
                CmbNonDeliveryReasonSorting.Items.Clear();                
                foreach (string items in PubConstClass.lstNonDeliveryList)
                {
                    string[] sArray = items.Split(',');
                    CmbNonDeliveryReasonSorting.Items.Add(sArray[0] + "：" + sArray[1]);                    
                }
                CmbNonDeliveryReasonSorting.SelectedIndex = 0;
                #endregion

                #region テスト通数
                CmbTestCount.Items.Clear();
                for(int i = 1; i < 100; i++)
                {
                    CmbTestCount.Items.Add(i.ToString("000"));
                }
                CmbTestCount.SelectedIndex = 4;
                #endregion

                // シリアルポートのオープン
                SerialPortQr.Open();
                LblError.Visible = false;
            }
            catch (Exception ex)
            {
                LblError.Text = ex.Message;
                LblError.Visible = true;
                MessageBox.Show(ex.Message, "【MainForm_Load】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「終了」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnEnd_Click(object sender, EventArgs e)
        {
            try
            {
                CommonModule.OutPutLogFile("シミュレーターの終了");
                this.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnEnd_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// データ受信
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void SerialPortQr_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string data;
            object[] args = new object[1];

            data = "";

            try
            {
                // シリアルポートをオープンしていない場合、処理を行わない。
                if (SerialPortQr.IsOpen == false)
                    return;
                // <CR>まで読み込む
                data = SerialPortQr.ReadTo("\r");

                // 受信データの格納
                BeginInvoke(new Delegate_RcvDataToTextBox(RcvDataToTextBox), data.ToString() + "\r");
            }
            catch (TimeoutException)
            {
                // ディスカードするデータ
                CommonModule.OutPutLogFile("■データ受信タイムアウトエラー：<CR>未受信で切り捨てたデータ：" + data);
            }
            catch (Exception ex)
            {
                CommonModule.OutPutLogFile("【SerialPortQr_DataReceived】" + ex.Message);
            }
        }

        /// <summary>
        /// 受信データによる各コマンド処理
        /// </summary>
        /// <param name="data">受信した文字列</param>
        /// <remarks></remarks>
        private void RcvDataToTextBox(string data)
        {
            string strMessage;

            try
            {
                CommonModule.OutPutLogFile("受信データ：" + data.Replace("\r", "<CR>"));

                LsbRecvBox.Items.Add(data.Replace("\r", "<CR>"));
                LsbRecvBox.SelectedIndex = LsbRecvBox.Items.Count - 1;

                // 受信データの先頭１文字を取得
                string sCommandType = data.Substring(0, 1);
                switch (sCommandType)
                {
                    case PubConstClass.CMD_RECIEVE_a:
                        // JOB設定内容

                        break;

                    case PubConstClass.CMD_RECIEVE_b:
                        // 開始コマンド
                        BtnStart.PerformClick();
                        break;

                    case PubConstClass.CMD_RECIEVE_c:
                        // 停止コマンド
                        BtnStop.PerformClick();
                        break;

                    default:
                        // 未定義コマンド
                        break;
                }
            }
            catch (Exception ex)
            {
                strMessage　 = "【RcvDataToTextBox】" + ex.Message;
                CommonModule.OutPutLogFile(strMessage);
                MessageBox.Show(strMessage, "システムエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「テストデータ送信」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSendTestData_Click(object sender, EventArgs e)
        {
            try
            {
                string sData = "";
                // ユニークキー（17桁）をセット
                if (TxtUniqueKey.Text.Trim() == "")
                {
                    sData += "     ";               // 物件ID（5桁）
                    sData += " ";                   // （1st/2st）
                    sData += "        ";            // 局出し日（YYYYMMDD）
                    sData += "                 ";   // ユニークキー（スペース：17桁）
                    sData += ",";
                }
                else
                {
                    //// 読取値（31桁）：物件ID（5桁）＋（1st/2st）＋局出し日（YYYYMMDD）＋ユニークキー（17桁）
                    //sData = TxtPropertyId.Text.Trim();                                          // 物件ID
                    //sData += "1";                                                               // （1st/2st）
                    //sData += dtTimPickPostalDate.Value.ToString("yyyyMMdd");                    // 局出し日（YYYYMMDD）
                    //sData += "-" + DateTime.Now.ToString("yyMMdd_");                             // ユニークキー（7桁）                    
                    ////sData += int.Parse(TxtUniqueKey.Text).ToString().PadLeft(CmbQRDigit.SelectedIndex + 9, '0') + ",";  // ユニークキー（xx桁）
                    ////sData += int.Parse(TxtUniqueKey.Text).ToString().PadLeft(CmbQRDigit.SelectedIndex + 9 - 7, '0') + "       ,";  // ユニークキー（xx桁）
                    //sData += TxtUniqueKey.Text + "       ,";  // ユニークキー（xx桁）

                    //// ユニークキーのインクリメント                        
                    ////TxtUniqueKey.Text = (int.Parse(TxtUniqueKey.Text) + 1).ToString().PadLeft(CmbQRDigit.SelectedIndex + 9, '0');

                    sData = TxtUniqueKey.Text + ",";
                    
                    TxtUniqueKey.Text = (int.Parse(TxtUniqueKey.Text) + 1).ToString();
                }

                // 判定（OK="0"/NG="1"）
                sData += CmbJudge.SelectedIndex.ToString("0") + ",";
                // エラーコード
                sData += CmbErrorCode.Text + ",";
                // 不着事由
                //sData += (CmbNonDeliveryReasonSorting.SelectedIndex + 1).ToString() + ",";
                // トレイ情報
                sData += CmbTray.Text + ",";

                // 先頭に「D,」を付加する
                sData = PubConstClass.CMD_SEND_D + "," + sData;
                // 送信データのセット                
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(sData + "\r");                
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add(sData);
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;
                CommonModule.OutPutLogFile($"送信データ：{sData}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnSendTestData_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「設定」アイコン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnMaintenance_Click(object sender, EventArgs e)
        {
            try
            {
                CommonModule.OutPutLogFile("「設定」アイコンクリック");
                MaintenanceForm form = new MaintenanceForm();
                form.Show(this);
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnMaintenance_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// メインフォームがアクティブになった時の処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Activated(object sender, EventArgs e)
        {
            try
            {
                this.Text = PubConstClass.pblMainFormTitle;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MainForm_Activated】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAutoSend_Click(object sender, EventArgs e)
        {
            try
            {
                if (TimSendData.Enabled)
                {
                    TimSendData.Enabled = false;
                }
                else
                {
                    TimSendData.Interval = 300;
                    TimSendData.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnAutoSend_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 連続送信処理用タイマー
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimSendData_Tick(object sender, EventArgs e)
        {
            BtnSendTestData.PerformClick();
        }

        /// <summary>
        /// 「開始」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnStart_Click(object sender, EventArgs e)
        {
            try
            {
                // 送信データのセット
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(PubConstClass.CMD_SEND_B + "\r");
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add($"{PubConstClass.CMD_SEND_B}<CR>");
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnStart_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「停止」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnStop_Click(object sender, EventArgs e)
        {
            try
            {
                // 送信データのセット
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(PubConstClass.CMD_SEND_C + "\r");
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add($"{PubConstClass.CMD_SEND_C}<CR>");
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnStop_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「エラー送信」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnError_Click(object sender, EventArgs e)
        {
            try
            {
                // 送信データのセット
                string sData = PubConstClass.CMD_SEND_E + "," + CmbErrorCode.Text;
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(sData + "\r");
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add($"{sData}<CR>");
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnError_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「確認」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnConfirmation_Click(object sender, EventArgs e)
        {
            try
            {
                // 送信データのセット
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(PubConstClass.CMD_SEND_A + "\r");
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add($"{PubConstClass.CMD_SEND_A}<CR>");
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnConfirmation_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int iIoIndex = 0;

        private void BtnInputOutput_Click(object sender, EventArgs e)
        {
            //string sSendData = "1234567890123456";

            try
            {
                string sData = "";
                sData += PubConstClass.CMD_SEND_K;

                if (iIoIndex < 10)
                {
                    sData += iIoIndex.ToString("0");
                }
                else
                {
                    sData += Convert.ToString(iIoIndex, 16).ToUpper();
                }                
                sData += DateTime.Now.ToString("HH:mm:ss.fff");

                // 送信データのセット
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(sData + "\r");
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add($"{sData}<CR>");
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;

                iIoIndex++;
                if (iIoIndex >= 16)
                {
                    iIoIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnConfirmation_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「I」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnICommand_Click(object sender, EventArgs e)
        {
            string sData = "";

            try
            {
                sData = PubConstClass.CMD_SEND_I + ",";
                sData += TxtICommand.Text;

                // 送信データのセット
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(sData + "\r");
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add($"{sData}<CR>");
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnICommand_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「J」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnJCommand_Click(object sender, EventArgs e)
        {
            string sData = "";
            try
            {
                sData = PubConstClass.CMD_SEND_J + ",";
                sData += TxtJCommand.Text;
                // 送信データのセット
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(sData + "\r");
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add($"{sData}<CR>");
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnJCommand_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「DIP-SW（T）」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDipSw_Click(object sender, EventArgs e)
        {
            try
            {
                // 送信データのセット
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(PubConstClass.CMD_SEND_T + "\r");
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add($"{PubConstClass.CMD_SEND_T}<CR>");
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnDipSw_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「QRデータ（L）」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnQrDataSend_Click(object sender, EventArgs e)
        {
            try
            {
                string sData = PubConstClass.CMD_SEND_L + ",";

                // 読取値（31桁）：物件ID（5桁）＋（1st/2st）＋局出し日（YYYYMMDD）＋ユニークキー（17桁）
                sData += TxtPropertyId.Text.Trim();                                     // 物件ID
                sData += "1";                                                           // （1st/2st）
                sData += dtTimPickPostalDate.Value.ToString("yyyyMMdd");                // 局出し日（YYYYMMDD）
                sData += "-" + DateTime.Now.ToString("yyMMdd_");                        // ユニークキー（8桁）
                //sData += int.Parse(TxtUniqueKey.Text).ToString("00000000000000000");    // ユニークキー（17桁）
                sData += int.Parse(TxtUniqueKey.Text).ToString().PadLeft(CmbQRDigit.SelectedIndex, '0') + ",";  // ユニークキー（17桁）

                // 送信データのセット
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(sData + "\r");
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add($"{sData}<CR>");
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnDipSw_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「自動テスト」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAutoTest_Click(object sender, EventArgs e)
        {
            try
            {
                int iCount = CmbTestCount.SelectedIndex + 1;

                for (int iTrayCount = 0; iTrayCount < 6; iTrayCount++) {

                    CmbTray.SelectedIndex = iTrayCount;

                    for (int iIndex = 0; iIndex < iCount; iIndex++) {
                        BtnQrDataSend.PerformClick();
                        // 100ミリ秒待機
                        Thread.Sleep(100);

                        BtnSendTestData.PerformClick();
                        // 700ミリ秒待機
                        Thread.Sleep(400);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnAutoTest_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「リセット」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnReset_Click(object sender, EventArgs e)
        {
            try
            {
                // 送信データのセット
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(PubConstClass.CMD_SEND_F + "\r");
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                LsbSendBox.Items.Add($"{PubConstClass.CMD_SEND_F}<CR>");
                LsbSendBox.SelectedIndex = LsbSendBox.Items.Count - 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnReset_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
