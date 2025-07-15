using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace QrSorterInspectionApp
{
    public partial class MenuForm : Form
    {
        private delegate void Delegate_RcvDataToTextBox(string data);

        public MenuForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// フォームロード処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuForm_Load(object sender, EventArgs e)
        {
            try
            {
                LblVersion.Text = PubConstClass.DEF_VERSION;
                CommonModule.OutPutLogFile("メニュー画面を表示しました");
                PubConstClass.pblOffLineMode = "0";
                // システム定義ファイル読込
                CommonModule.ReadSystemDefinition();
                // エラーメッセージファイル読込
                PubConstClass.dicErrorCodeData = new Dictionary<string, string>();
                CommonModule.ReadErrorMessageFile();
                // 読取機能項目ファイル読込
                CommonModule.ReadReadFunctionItemFile();

                #region シリアルポートの設定とオープン
                // データ受信イベントの設定
                SerialPort.DataReceived += new SerialDataReceivedEventHandler(SerialPort_DataReceived);
                // シリアルポート名の設定
                SerialPort.PortName = PubConstClass.pblComPort;
                // シリアルポートの通信速度指定
                switch (PubConstClass.pblComSpeed)
                {
                    case "0":
                        SerialPort.BaudRate = 4800;
                        break;
                    case "1":
                        SerialPort.BaudRate = 9600;
                        break;
                    case "2":
                        SerialPort.BaudRate = 19200;
                        break;
                    case "3":
                        SerialPort.BaudRate = 38400;
                        break;
                    case "4":
                        SerialPort.BaudRate = 57600;
                        break;
                    case "5":
                        SerialPort.BaudRate = 115200;
                        break;
                    default:
                        SerialPort.BaudRate = 38400;
                        break;
                }
                // シリアルポートのパリティ指定
                switch (PubConstClass.pblComParityVar)
                {
                    case "0":
                        SerialPort.Parity = Parity.Odd;
                        break;
                    case "1":
                        SerialPort.Parity = Parity.Even;
                        break;
                    default:
                        SerialPort.Parity = Parity.Even;
                        break;
                }

                // シリアルポートのパリティ有無
                if (PubConstClass.pblComIsParity == "0")
                    SerialPort.Parity = Parity.None;

                // シリアルポートのビット数指定
                switch (PubConstClass.pblComDataLength)
                {
                    case "0":
                        SerialPort.DataBits = 8;
                        break;
                    case "1":
                        SerialPort.DataBits = 7;
                        break;
                    default:
                        SerialPort.DataBits = 8;
                        break;
                }

                // シリアルポートのストップビット指定
                switch (PubConstClass.pblComStopBit)
                {
                    case "0":
                        SerialPort.StopBits = StopBits.One;
                        break;
                    case "1":
                        SerialPort.StopBits = StopBits.Two;
                        break;
                    default:
                        SerialPort.StopBits = StopBits.One;
                        break;
                }

                // シリアルポートのオープン
                SerialPort.Open();
                #endregion

                // シリアルポートにデータ送信（動作可コマンド）
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(PubConstClass.CMD_SEND_a + "\r");
                SerialPort.Write(dat, 0, dat.GetLength(0));
                CommonModule.OutPutLogFile($"〓【メニュー画面】送信データ：{PubConstClass.CMD_SEND_a}");

                // ディスクの空き領域をチェック
                CommonModule.CheckAvairableFreeSpace();

                LblUserInfo.Text = "";
                LblUserInfo.Text += "ＩＤ：" + PubConstClass.sUserId + Environment.NewLine;
                LblUserInfo.Text += "名前：" + PubConstClass.sUserName + Environment.NewLine;
                LblUserInfo.Text += "権限：" + PubConstClass.sUserAuthority + Environment.NewLine;
                LblUserInfo.Text += "ＰＷ：" + PubConstClass.sUserPassword;

                if (PubConstClass.sUserAuthority == "OP")
                {
                    BtnSetting.Enabled = false;
                    BtnMaintenance.Enabled = false;
                }

                LblStatus.Visible = false;

                // DIP-SW 情報送信
                MyProcDipSw();

                // 現在の日付（年月日）を求める
                DateTime dtCurrent = DateTime.Now;
                int intMinusMonth = Convert.ToInt32(PubConstClass.pblSaveLogMonth);
                // 現在日付から１ヶ月を減算
                DateTime dtPassDate = dtCurrent.AddMonths(-(intMinusMonth));
                CommonModule.OutPutLogFile($"現在の日付から{intMinusMonth}ヶ月より前（{dtPassDate}）のファイルを削除します。");
                CommonModule.DeleteLogFiles(intMinusMonth);
            }
            catch (Exception ex)
            {
                LblStatus.Text = ex.Message;
                LblStatus.Visible = true;
                MessageBox.Show(ex.Message, "【MenuForm_Load】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「ログアウト」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLogOut_Click(object sender, EventArgs e)
        {
            if (SerialPort.IsOpen)
            {
                SerialPort.Close();
            }
            CommonModule.OutPutLogFile("メニュー画面：「ログアウト」ボタンクリック");
            Owner.Show();
            this.Dispose();
        }

        /// <summary>
        /// 「終了」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnEnd_Click(object sender, EventArgs e)
        {
            CommonModule.OutPutLogFile("メニュー画面：「終了」ボタンクリック");
            Owner.Dispose();
            this.Dispose();
        }

        /// <summary>
        /// 「設定」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSetting_Click(object sender, EventArgs e)
        {
            try
            {
                CommonModule.OutPutLogFile("メニュー画面：「設定」ボタンクリック");
                PubConstClass.sJobFileNameFromInspectionForm = "";
                SettingForm form = new SettingForm();
                form.Show(this);
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnSetting_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「QRソーター検査」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnQrSorterInspect_Click(object sender, EventArgs e)
        {
            try
            {
                if (SerialPort.IsOpen)
                {
                    SerialPort.Close();
                }

                CommonModule.OutPutLogFile("メニュー画面：「QRソーター検査」ボタンクリック");
                QrSorterInspectionForm form = new QrSorterInspectionForm();
                form.ShowDialog();

                if (!SerialPort.IsOpen)
                {
                    SerialPort.Open();
                }
                LblStatus.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnQrSorterInspect_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「SV・OP設定」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAccountSet_Click(object sender, EventArgs e)
        {
            try
            {
                CommonModule.OutPutLogFile("メニュー画面：「SV・OP設定」ボタンクリック");
                RegisterAccountForm form = new RegisterAccountForm();
                form.Show(this);
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnAccountSet_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「ログ管理」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLogManagement_Click(object sender, EventArgs e)
        {
            try
            {
                CommonModule.OutPutLogFile("メニュー画面：「ログ管理」ボタンクリック");
                LogListForm form = new LogListForm();
                form.Show(this);
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnLogManagement_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「保守」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnMaintenance_Click(object sender, EventArgs e)
        {
            try
            {
                if (SerialPort.IsOpen)
                {
                    SerialPort.Close();
                }

                CommonModule.OutPutLogFile("メニュー画面：「保守」ボタンクリック");
                MaintenanceForm form = new MaintenanceForm();
                form.ShowDialog();

                if (!SerialPort.IsOpen)
                {
                    SerialPort.Open();
                }
                LblStatus.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnLogManagement_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// バージョン表示ラベルのダブルクリック処理（デバッグ用）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LblVersion_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                if (LblUserInfo.Visible == false)
                {
                    LblUserInfo.Visible = true;
                }
                else
                {
                    LblUserInfo.Visible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【LblVersion_DoubleClick】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 検査装置からのデータ受信処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string data;
            object[] args = new object[1];

            data = "";

            try
            {
                // シリアルポートをオープンしていない場合、処理を行わない。
                if (SerialPort.IsOpen == false)
                    return;
                // <CR>まで読み込む
                data = SerialPort.ReadTo("\r");
                // 受信データの格納
                BeginInvoke(new Delegate_RcvDataToTextBox(RcvDataToTextBox), data.ToString() + "\r");
            }
            catch (TimeoutException)
            {
                // ディスカードするデータ
                CommonModule.OutPutLogFile("データ受信タイムアウトエラー：<CR>未受信で切り捨てたデータ：" + data);
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
                CommonModule.OutPutLogFile($"■【メニュー画面】受信データ：{data.Replace("\r", "<CR>")}");

                // 受信データの先頭１文字を取得
                string sCommandType = data.Substring(0, 1);
                switch (sCommandType)
                {
                    case PubConstClass.CMD_RECIEVE_A:
                    case PubConstClass.CMD_RECIEVE_B:
                    case PubConstClass.CMD_RECIEVE_D:
                    case PubConstClass.CMD_RECIEVE_L:
                        // 検査不可コマンドの送信
                        SendSerialData(PubConstClass.CMD_SEND_e);
                        break;

                    case PubConstClass.CMD_RECIEVE_C:
                    case PubConstClass.CMD_RECIEVE_K:
                    case PubConstClass.CMD_RECIEVE_I:
                    case PubConstClass.CMD_RECIEVE_J:
                        // コマンドを無視する
                        break;

                    case PubConstClass.CMD_RECIEVE_E:
                        // エラーコマンド
                        MyProcError(data);
                        break;

                    case PubConstClass.CMD_RECIEVE_F:
                        // エラーリセットコマンド
                        MyProcErrorReset();
                        break;

                    case PubConstClass.CMD_RECIEVE_T:
                        // DIP-SW 情報送信
                        MyProcDipSw();
                        break;

                    default:
                        // 未定義コマンド
                        CommonModule.OutPutLogFile($"【メニュー画面】未定義コマンドです：{data.Replace("\r", "<CR>")}");
                        break;
                }
            }
            catch (Exception ex)
            {
                strMessage = "【RcvDataToTextBox】" + ex.Message;
                CommonModule.OutPutLogFile(strMessage);
                MessageBox.Show(strMessage, "システムエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// シリアルデータ送信処理
        /// </summary>
        /// <param name="sData"></param>
        private void SendSerialData(string sData)
        {
            try
            {
                // 送信データのセット
                byte[] dat = Encoding.GetEncoding("SHIFT-JIS").GetBytes(sData + "\r");
                SerialPort.Write(dat, 0, dat.GetLength(0));
                CommonModule.OutPutLogFile($"〓【メニュー画面】送信データ：{sData}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【SendSerialData】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// DIP-SW情報の送信
        /// </summary>
        private void MyProcDipSw()
        {
            string sData;

            try
            {
                sData = PubConstClass.CMD_SEND_t + "," + PubConstClass.pblDipSw;
                // シリアルデータ送信
                SendSerialData(sData);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyProcQrData】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CommonModule.OutPutLogFile("【MyProcQrData】" + ex.Message);
            }
        }

        /// <summary>
        /// エラーコマンド処理
        /// </summary>
        /// <param name="sData"></param>
        private void MyProcError(string sData)
        {
            string sErrorCode;

            try
            {
                sErrorCode = sData.Substring(2, 3);
                if (PubConstClass.dicErrorCodeData.ContainsKey(sErrorCode))
                {
                    // 存在する場合
                    LblStatus.Text = $"エラーNo.{sErrorCode}（{PubConstClass.dicErrorCodeData[sErrorCode].Replace(",", "）")}";
                    CommonModule.OutPutLogFile($"エラーNo.{sErrorCode}（{PubConstClass.dicErrorCodeData[sErrorCode].Replace(",", "）")}");
                }
                else
                {
                    LblStatus.Text = $"エラーNo.{sErrorCode}（未定義エラー）未定義のエラー番号です";
                    CommonModule.OutPutLogFile($"エラーNo.{sErrorCode}（未定義エラー）未定義のエラー番号です");
                }
                LblStatus.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyProcError】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// エラーリセットコマンド
        /// </summary>
        private void MyProcErrorReset()
        {
            try
            {
                LblStatus.Text = "";
                LblStatus.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyProcErrorReset】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
