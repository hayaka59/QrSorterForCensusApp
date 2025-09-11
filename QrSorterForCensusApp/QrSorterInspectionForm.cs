using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QrSorterInspectionApp
{
    public partial class QrSorterInspectionForm : Form
    {
        private delegate void Delegate_RcvDataToTextBox(string data);

        private List<string> lstPastReceivedQrData = new List<string>();

        private string sDateOfReceipt;          // 受領日
        private bool   bDateOfReceipt;          // 受領日入力
        private string sNonDeliveryReason1;     // 不着事由１
        private string sNonDeliveryReason2;     // 不着事由２
        private int    iStatus;                 // 検査中ステータス
        private bool   bIsDuplicateCheck;       // 重複チェック
        private bool   bIsJobChange;            // JOB変更フラグ
                
        private int iOKCount = 0;               // OK用カウンタ
        private int iNGCount = 0;               // NG用カウンタ
        private int iBox1Count = 0;             // ボックス１用カウンタ
        private int iBox2Count = 0;             // ボックス２用カウンタ
        private int iBox3Count = 0;             // ボックス３用カウンタ
        private int iBox4Count = 0;             // ボックス４用カウンタ
        private int iBox5Count = 0;             // ボックス５用カウンタ
        private int iBoxECount = 0;             // ボックス（Eject）用カウンタ
        private int intOkSesanCounter = 0;      // 処理数No.カウンタ
        private int intNgSesanCounter = 0;      // 処理数No.カウンタ

        private string sJobFolderName;          // JOBフォルダ名       
        private string sFolderName1;            // グループ１フォルダ名
        private string sFolderName2;            // グループ２フォルダ名
        private string sFolderName3;            // グループ３フォルダ名
        private string sFolderName4;            // グループ４フォルダ名
        private string sFolderName5;            // グループ５フォルダ名        
        private string sFileNameForGroup1;      // グループ１操作ログファイル名
        private string sFileNameForGroup2;      // グループ２操作ログファイル名
        private string sFileNameForGroup3;      // グループ３操作ログファイル名
        private string sFileNameForGroup4;      // グループ４操作ログファイル名
        private string sFileNameForGroup5;      // グループ５操作ログファイル名

        #region ログ保存関係
        private string sProcessingModeName;     // 処理モード名
        private string sProcessingDate;         // 処理日
        private string sBoxLabelNumber;         // 箱ラベル番号
        private string sInquiryNumber;          // 問い合わせ番号 
        private string sReceiptDate;            // 受領日
        private string sFolderNameForOkLog;     // OK用の操作ログ格納フォルダ名
        private string sFolderNameForAllLog;    // 全件用の操作ログ格納フォルダ名
        private string sFolderNameForErrorLog;  // エラーログ格納フォルダ名
        private string sFileNameForOkLog;       // OK用の操作ログファイル名
        private string sFileNameForAllLog;      // 全件用の操作ログファイル名
        private string sFileNameForErrorLog;    // エラーログファイル名
        #endregion
        private bool bManualEntryFlg = false;   // 手動登録中フラグ

        private byte[] buffer = new byte[1024];
        private int bufferIndex = 0;

        public QrSorterInspectionForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// フォームロード処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void QrSorterInspectionForm_Load(object sender, EventArgs e)
        {
            try
            {
                LblVersion.Text = PubConstClass.DEF_VERSION;
                CommonModule.OutPutLogFile("QRソーター検査画面を表示しました");
                // 年月日時分秒タイマーセット
                TimDateTime.Interval = 1000;
                TimDateTime.Enabled = true;

                CmbMode.Items.Clear();
                CmbMode.Items.Add("受付モード");
                CmbMode.Items.Add("箱詰めモード");
                CmbMode.SelectedIndex = 1;

                LblOffLine.Text = "箱詰めモード";
                TxtBoxLabelNumber.Enabled = true;
                TxtInquiryNumber.Enabled = true;

                // 受領日
                sReceiptDate = DtpDateReceipt.Value.ToString("yyyyMMdd");

                #region OK履歴のヘッダー設定
                // ListViewのカラムヘッダー設定
                LsvOKHistory.View = View.Details;                
                ColumnHeader colOK1 = new ColumnHeader();
                ColumnHeader colOK2 = new ColumnHeader();
                ColumnHeader colOK3 = new ColumnHeader();
                ColumnHeader colOK4 = new ColumnHeader();
                ColumnHeader colOK5 = new ColumnHeader();
                colOK1.Text = "No.";
                colOK2.Text = "日時";
                colOK3.Text = "読取値";
                colOK4.Text = "判定";
                colOK5.Text = "トレイ";
                colOK1.TextAlign = HorizontalAlignment.Center;
                colOK2.TextAlign = HorizontalAlignment.Center;
                colOK3.TextAlign = HorizontalAlignment.Center;
                colOK4.TextAlign = HorizontalAlignment.Center;
                colOK5.TextAlign = HorizontalAlignment.Center;
                colOK1.Width = 75;          // 
                colOK2.Width = 200;         // 
                colOK3.Width = 370;         // 
                colOK4.Width = 75;          // 
                colOK5.Width = 70;          // 
                ColumnHeader[] colHeaderOK = new[] { colOK1, colOK2, colOK3, colOK4, colOK5 };                
                LsvOKHistory.Columns.AddRange(colHeaderOK);
                #endregion
                #region NG履歴のヘッダー設定
                LsvNGHistory.View = View.Details;
                ColumnHeader colNG1 = new ColumnHeader();
                ColumnHeader colNG2 = new ColumnHeader();
                ColumnHeader colNG3 = new ColumnHeader();
                ColumnHeader colNG4 = new ColumnHeader();
                ColumnHeader colNG5 = new ColumnHeader();
                colNG1.Text = "No.";
                colNG2.Text = "日時";
                colNG3.Text = "読取値";
                colNG4.Text = "判定";
                colNG5.Text = "トレイ";
                colNG1.TextAlign = HorizontalAlignment.Center;
                colNG2.TextAlign = HorizontalAlignment.Center;
                colNG3.TextAlign = HorizontalAlignment.Center;
                colNG4.TextAlign = HorizontalAlignment.Center;
                colNG5.TextAlign = HorizontalAlignment.Center;
                colNG1.Width = 75;          // 
                colNG2.Width = 200;         // 
                colNG3.Width = 370;         // 
                colNG4.Width = 75;          // 
                colNG5.Width = 70;          // 
                ColumnHeader[] colHeaderNG = new[] { colNG1, colNG2, colNG3, colNG4, colNG5 };
                LsvNGHistory.Columns.AddRange(colHeaderNG);
                #endregion
                #region 不着事由区分
                CommonModule.ReadNonDeliveryList();
                CmbNonDeliveryReasonSorting1.Items.Clear();
                CmbNonDeliveryReasonSorting2.Items.Clear();
                foreach (string items in PubConstClass.lstNonDeliveryList)
                {
                    string[] sArray = items.Split(',');
                    CmbNonDeliveryReasonSorting1.Items.Add(sArray[0] + "：" + sArray[1]);
                    CmbNonDeliveryReasonSorting2.Items.Add(sArray[0] + "：" + sArray[1]);
                }
                CmbNonDeliveryReasonSorting1.SelectedIndex = 0;
                CmbNonDeliveryReasonSorting2.SelectedIndex = 0;
                #endregion                                
                #region QRフィーダーカウンタクリア
                LblTotalCount.Text = "0";   // 総数カウンタクリア
                LblOKCount.Text = "0";      // OKカウンタクリア
                LblNGCount.Text = "0";      // NGカウンタクリア
                #endregion
                #region ソーターポケットカウンタクリア
                LblBox1.Text = "0";
                LblBox2.Text = "0";
                LblBox3.Text = "0";
                LblBox4.Text = "0";
                LblBox5.Text = "0";
                LblBoxEject.Text = "0";
                #endregion
                #region ソーターポケット予測値クリア
                LblPocket1.Text = "";
                LblPocket2.Text = "";
                LblPocket3.Text = "";
                LblPocket4.Text = "";
                LblPocket5.Text = "";
                LblPocketEject.Text = "";
                #endregion
                #region ソーターポケットタイトルクリア
                LblBoxTitle1.Text = "";
                LblBoxTitle2.Text = "";
                LblBoxTitle3.Text = "";
                LblBoxTitle4.Text = "";
                LblBoxTitle5.Text = "";
                #endregion
                #region ソーターポケット数量クリア
                LblQuantity1.Text = "---";
                LblQuantity2.Text = "---";
                LblQuantity3.Text = "---";
                LblQuantity4.Text = "---";
                LblQuantity5.Text = "---";
                #endregion

                LblQrReadData.Text = "";
                bIsJobChange = false;

                LstSettingInfomation.Items.Clear();
                LstSettingInfomation.Items.Add("【設定内容】");
                LstSettingInfomation.Items.Add("Ｗフィード検査：");
                //LstSettingInfomation.Items.Add("超音波検査　　：");
                LstSettingInfomation.Items.Add("桁数チェック　：");
                LstSettingInfomation.Items.Add("読取機能　　　：");
                LstSettingInfomation.Items.Add("読取チェック　：");
                LstSettingInfomation.Items.Add("読取位置　　　：");
                //LstSettingInfomation.Items.Add("C/D チェック　：");

                // 過去に受信したQRデータ一覧のクリア
                lstPastReceivedQrData.Clear();
                LblDuplicateCheck.Text = "重複チェック";
                PubConstClass.sPrevDtpDateReceipt = "";  // 前回の受領日
                PubConstClass.sPrevNonDelivery1 = "";    // 前回の不着事由仕分け１
                PubConstClass.sPrevNonDelivery2 = "";    // 前回の不着事由仕分け２

                // yyyyMMdd HHmmss形式で現在日時を取得
                // エラーログ用のログ保存用フォルダの作成
                sFolderNameForErrorLog = CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder) +
                                         "エラーログ\\" + DateTime.Now.ToString("yyyyMMdd");
                if (Directory.Exists(sFolderNameForErrorLog) == false)
                {
                    Directory.CreateDirectory(sFolderNameForErrorLog);
                }
                sFileNameForErrorLog = $"\\国勢調査用_errorlog_{DateTime.Now.ToString("yyyyMMdd")}_" +
                                       $"{DateTime.Now.ToString("yyyyMMdd")}121234.csv";

                //string sOutPutDateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                //sFileNameForOkLog = $"{sFolderNameForOkLog}\\uketuke_{PubConstClass.pblMachineName}_{sReceiptDate}_{sOutPutDateTime}.csv";
                //sFileNameForAllLog = $"{sFolderNameForAllLog}\\uketuke_{PubConstClass.pblMachineName}_{sReceiptDate}_{sOutPutDateTime}（全件）.csv";

                // 停止中
                SetStatus(0);
                // JOB選択ラベルクリア
                LblSelectedFile.Text = "";
                // 「検査開始」ボタン使用不可
                BtnStartInspection.Enabled = false;
                // 「設定」ボタン使用不可
                BtnSetting.Enabled = false;

                #region シリアルポートの設定
                // データ受信イベントの設定
                SerialPortQr.DataReceived += new SerialDataReceivedEventHandler(SerialPortQr_DataReceived);
                // シリアルポート名の設定
                SerialPortQr.PortName = PubConstClass.pblComPort;
                // シリアルポートの通信速度指定
                switch (PubConstClass.pblComSpeed)
                {
                    case "0":
                        {
                            SerialPortQr.BaudRate = 4800;
                            break;
                        }

                    case "1":
                        {
                            SerialPortQr.BaudRate = 9600;
                            break;
                        }

                    case "2":
                        {
                            SerialPortQr.BaudRate = 19200;
                            break;
                        }

                    case "3":
                        {
                            SerialPortQr.BaudRate = 38400;
                            break;
                        }

                    case "4":
                        {
                            SerialPortQr.BaudRate = 57600;
                            break;
                        }

                    case "5":
                        {
                            SerialPortQr.BaudRate = 115200;
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
                            break;
                        }

                    case "1":
                        {
                            SerialPortQr.Parity = Parity.Even;
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
                    SerialPortQr.Parity = Parity.None;
                // シリアルポートのビット数指定
                switch (PubConstClass.pblComDataLength)
                {
                    case "0":
                        {
                            SerialPortQr.DataBits = 8;
                            break;
                        }

                    case "1":
                        {
                            SerialPortQr.DataBits = 7;
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
                            break;
                        }

                    case "1":
                        {
                            SerialPortQr.StopBits = StopBits.Two;
                            break;
                        }

                    default:
                        {
                            SerialPortQr.StopBits = StopBits.One;
                            break;
                        }
                }
                #endregion
                // シリアルポートのオープン
                SerialPortQr.Open();
                LblError.Visible = false;

                // リストビューのダブルバッファを有効とする
                EnableDoubleBuffering(LsvOKHistory);
                EnableDoubleBuffering(LsvNGHistory);

                if (PubConstClass.pblOffLineMode == "1")
                {
                    LblTitle.Text = "国勢調査用アプリ検査画面";
                    LblOffLine.Visible = true;
                    // ボックス１～５のカウンタの表示フォントを変更
                    LblBox1.Font = new Font("メイリオ", 28);
                    LblBox2.Font = new Font("メイリオ", 28);
                    LblBox3.Font = new Font("メイリオ", 28);
                    LblBox4.Font = new Font("メイリオ", 28);
                    LblBox5.Font = new Font("メイリオ", 28);
                }
                else
                {
                    LblTitle.Text = "QRフィーダー＆ソーター検査画面";
                    LblOffLine.Visible = false;
                    // ボックス１～５のカウンタの表示フォントを変更
                    LblBox1.Font = new Font("メイリオ", 48);
                    LblBox2.Font = new Font("メイリオ", 48);
                    LblBox3.Font = new Font("メイリオ", 48);
                    LblBox4.Font = new Font("メイリオ", 48);
                    LblBox5.Font = new Font("メイリオ", 48);
                }

                //　固定のジョブファイル（国勢調査用JOB設定.csv）の読込処理
                LoadingFixedJobFile();

                // 箱ラベル番号入力にフォーカスを当てる
                TxtBoxLabelNumber.Focus();

                TxtQrReadData.Enabled = false;
            }
            catch (Exception ex)
            {
                LblError.Text = ex.Message;
                LblError.Visible = true;
                MessageBox.Show(ex.Message, "【QrSorterInspectionForm_Load】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 検査ログフォルダの作成と検査ログファイル名の確保
        /// </summary>
        private void CreateInspectionLogFolder()
        {
            try
            {
                if (LblSelectedFile.Text.Trim()=="")
                {
                    CommonModule.OutPutLogFile("JOB未選択状態で CreateInspectionLogFolder() が呼ばれました");
                    return;
                }

                // 箱ラベル番号
                sBoxLabelNumber = TxtBoxLabelNumber.Text.Trim();
                // 問い合わせ番号
                sInquiryNumber = TxtInquiryNumber.Text.Trim();
                // 受領日
                sReceiptDate = DtpDateReceipt.Value.ToString("yyyyMMdd");

                //string sModeName;
                if (CmbMode.SelectedIndex == 0)
                {
                    // 受付モード
                    sProcessingModeName = "受付用";
                }
                else
                {
                    // 箱詰めモード
                    sProcessingModeName = "箱詰め用";
                }
                // 処理日の取得
                sProcessingDate = DateTime.Now.ToString("yyyyMMdd");

                // ＯＫ用の検査ログ保存用フォルダの作成
                sFolderNameForOkLog = CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder) +
                                      sProcessingModeName + "\\" + sProcessingDate;
                if (Directory.Exists(sFolderNameForOkLog) == false)
                {
                    Directory.CreateDirectory(sFolderNameForOkLog);
                }

                // 全件用の検査ログ保存用フォルダの作成
                sFolderNameForAllLog = CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder) +
                                       "受付・箱詰め用\\" + sProcessingDate;
                if (Directory.Exists(sFolderNameForAllLog) == false)
                {
                    Directory.CreateDirectory(sFolderNameForAllLog);
                }

                //// エラーログ用のログ保存用フォルダの作成
                //sFolderNameForErrorLog = CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder) +
                //                         "エラーログ\\" + sProcessingDate;
                //if (Directory.Exists(sFolderNameForErrorLog) == false)
                //{
                //    Directory.CreateDirectory(sFolderNameForErrorLog);
                //}

                string sOutPutDateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                string sOutPutDate = DateTime.Now.ToString("yyyyMMdd");

                if (CmbMode.SelectedIndex == 0)
                {
                    // 受付モード
                    if (sFileNameForOkLog == "")
                    {
                        sFileNameForOkLog = $"{sFolderNameForOkLog}\\uketuke_{PubConstClass.pblMachineName}_{sReceiptDate}_{sOutPutDateTime}.csv";
                        sFileNameForAllLog = $"{sFolderNameForAllLog}\\uketuke_{PubConstClass.pblMachineName}_{sReceiptDate}_{sOutPutDateTime}（全件）.csv";
                    }
                }
                else
                {
                    // 箱詰めモード
                    sFileNameForOkLog = $"{sFolderNameForOkLog}\\{sBoxLabelNumber}_{sInquiryNumber}_{sReceiptDate}_{sOutPutDateTime}.csv";
                    sFileNameForAllLog = $"{sFolderNameForAllLog}\\{sBoxLabelNumber}_{sInquiryNumber}_{sReceiptDate}_{sOutPutDateTime}（全件）.csv";
                }

                sFolderNameForErrorLog = CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder) +
                                       "エラーログ\\" + sOutPutDate;
                sFileNameForErrorLog = $"国勢調査用_errorlog_{sOutPutDate}_{sOutPutDateTime}.csv";

                LblFdrInfo1.Text = sFolderNameForOkLog;
                LblGrpInfo1.Text = sFileNameForOkLog;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace, "【CreateInspectionLogFolder】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// コントロールのDoubleBufferedプロパティをTrueにする
        /// </summary>
        /// <param name="control"></param>
        private void EnableDoubleBuffering(Control control)
        {
            control.GetType().InvokeMember("DoubleBuffered",
                                            BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
                                            null/* TODO Change to default(_) if this is not a reference type */,
                                            control,
                                            new object[] { true }
                                            );
        }

        /// <summary>
        /// 「戻る」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnClose_Click(object sender, EventArgs e)
        {
            try
            {
                if (LblSelectedFile.Text.Trim() != "")
                {
                    DialogResult dialogResult= MessageBox.Show("メニュー画面に戻りますか？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Cancel) {
                        // キャンセル
                        return;
                    }
                }
                // メニュー画面へ戻る
                this.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnClose_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 現在日付と時刻の表示
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimDateTime_Tick(object sender, EventArgs e)
        {
            int iBoxCount;

            try
            {
                // 現在時刻の表示
                LblDateTime.Text = DateTime.Now.ToString("yyyy年MM月dd日(ddd) HH:mm:ss");

                if (bIsReset)
                {
                    bIsReset = false;
                    // シリアルデータ送信
                    SendSerialData(PubConstClass.CMD_SEND_d);
                    LblError.Visible = false;
                    // 停止中
                    SetStatus(0);
                }

                iBoxCount = int.Parse(LblBox1.Text);
                //if (int.Parse(LblBox1.Text) >= 850)
                if (iBoxCount >= 850)
                {
                    if (LblOffLine.BackColor == Color.Yellow)
                    {
                        LblOffLine.BackColor = Color.WhiteSmoke;
                    }
                    else
                    {
                        LblOffLine.BackColor = Color.Yellow;
                    }
                }
                else
                {
                    LblOffLine.BackColor = Color.WhiteSmoke;
                }

                //// 900セット以上の時は、50で割り切れるかをチェックする
                //if (iBoxCount >= 200) // 900
                //{
                //    if (iBox1Count % 50 == 0)
                //    {
                //        LblOffLine.BackColor = Color.WhiteSmoke;
                //        // 900、950、1000、1050、110、、と50単位で停止する。
                //        // シリアルデータ送信
                //        SendSerialData(PubConstClass.CMD_SEND_c);
                //        LblError.Visible = false;

                //        MyProcStop();
                //    }
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【TimDateTime_Tick】", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                SerialPortQr.Write(dat, 0, dat.GetLength(0));
                CommonModule.OutPutLogFile($"〓送信データ：{sData}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【SendSerialData】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「検査開始」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnStartInspection_Click(object sender, EventArgs e)
        {
            try
            {
                // 各入力フィールドの桁数チェックを行う
                if (CheckNumberOfDigits())
                {
                    // シリアルデータ送信
                    SendSerialData(PubConstClass.CMD_SEND_b);
                    // 検査開始時のチェック
                    CheckStartUp();
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "【BtnStartInspection_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private bool CheckNumberOfDigits()
        {
            bool bRetVal = true;
            try
            {
                if (CmbMode.SelectedIndex == 1)
                {
                    // 箱詰めモードの場合のみチェックする
                    // 箱ラベル番号の桁数チェック
                    if (TxtBoxLabelNumber.Text.Trim().Length != 18)
                    {
                        bRetVal = false;
                        //「検査終了」とする
                        // シリアルデータ送信
                        SendSerialData(PubConstClass.CMD_SEND_c);
                        LblError.Visible = false;
                        // 18桁でない場合
                        MessageBox.Show("箱ラベル番号は、18桁で入力して下さい", "確認", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    // 問い合わせ番号の桁数チェック
                    if (TxtInquiryNumber.Text.Trim().Length != 12)
                    {
                        bRetVal = false;
                        //「検査終了」とする
                        // シリアルデータ送信
                        SendSerialData(PubConstClass.CMD_SEND_c);
                        LblError.Visible = false;
                        // 11～13桁でない場合
                        MessageBox.Show("問い合わせ番号は、12桁で入力して下さい", "確認", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                // 読み取り値の桁数チェック
                if (!(TxtCheckReading.Text.Trim() == "" || 
                      TxtCheckReading.Text.Trim().Length == 2 || 
                      TxtCheckReading.Text.Trim().Length == 5))
                {
                    bRetVal = false;
                    // シリアルデータ送信
                    SendSerialData(PubConstClass.CMD_SEND_c);
                    LblError.Visible = false;
                    // 2桁か5桁でない場合
                    MessageBox.Show("読み取り値は、空白か、2桁か 5桁に設定して下さい", "確認", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                return bRetVal;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【CheckNumberOfDigits】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 検査開始時のチェック
        /// </summary>
        private void CheckStartUp()
        {
            try
            {
                // エラー状況の非表示
                LblError.Visible = false;

                if (PubConstClass.sPrevDtpDateReceipt == "")
                {
                    // １回目の検査開始処理
                    CreateInspectionLogFolder();
                    // シリアルデータ送信（JOB選択）
                    SendSerialData(PubConstClass.CMD_SEND_h);
                }
                else
                {
                    if (bIsJobChange ||
                        PubConstClass.sPrevDtpDateReceipt != DtpDateReceipt.Text ||
                        PubConstClass.sPrevNonDelivery1 != CmbNonDeliveryReasonSorting1.Text ||
                        PubConstClass.sPrevNonDelivery2 != CmbNonDeliveryReasonSorting2.Text)
                    {
                        // JOBが変更された、受領日または不着事由仕分け１、２が変更された。
                        //MessageBox.Show("JOB設定が変更されました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        CommonModule.OutPutLogFile($"【JOB設定が変更されました】" +
                                                    $"JOB変更フラグ＝{bIsJobChange}／" +
                                                    $"受領日＝{DtpDateReceipt.Text}／" +
                                                    $"仕分け１＝{CmbNonDeliveryReasonSorting1.Text}／" +
                                                    $"仕分け２＝{CmbNonDeliveryReasonSorting2.Text}");
                        CreateInspectionLogFolder();
                        // シリアルデータ送信（JOB選択）
                        SendSerialData(PubConstClass.CMD_SEND_h);

                        // 各表示カウンタクリア
                        LblTotalCount.Text = "0";
                        LblOKCount.Text = "0";
                        LblNGCount.Text = "0";
                        // ポケット１～５の表示カウンタクリア
                        LblBox1.Text = "0";
                        LblBox2.Text = "0";
                        LblBox3.Text = "0";
                        LblBox4.Text = "0";
                        LblBox5.Text = "0";
                        LblBoxEject.Text = "0";
                        // 内部カウンタのクリア
                        iOKCount = 0;               // OK用カウンタ
                        iNGCount = 0;               // NG用カウンタ
                        iBox1Count = 0;             // ボックス１用カウンタ
                        iBox2Count = 0;             // ボックス２用カウンタ
                        iBox3Count = 0;             // ボックス３用カウンタ
                        iBox4Count = 0;             // ボックス４用カウンタ
                        iBox5Count = 0;             // ボックス５用カウンタ
                        iBoxECount = 0;             // ボックス（Eject）用カウンタ
                        intOkSesanCounter = 0;      // OK処理数No.カウンタ
                        intNgSesanCounter = 0;      // NG処理数No.カウンタ
                                                    // 受信データ表示領域のクリア
                        LblPocket1.Text = "";
                        LblPocket2.Text = "";
                        LblPocket3.Text = "";
                        LblPocket4.Text = "";
                        LblPocket5.Text = "";
                        LblPocketEject.Text = "";
                        // OK履歴とNG履歴のクリア
                        LsvOKHistory.Items.Clear();
                        LsvNGHistory.Items.Clear();

                        // 過去に受信したQRデータ一覧のクリア
                        lstPastReceivedQrData.Clear();
                    }
                }
                bIsJobChange = false;
                // 設定値の保存
                PubConstClass.sPrevDtpDateReceipt = DtpDateReceipt.Text;                // 前回の受領日
                PubConstClass.sPrevNonDelivery1 = CmbNonDeliveryReasonSorting1.Text;    // 前回の不着事由仕分け１
                PubConstClass.sPrevNonDelivery2 = CmbNonDeliveryReasonSorting2.Text;    // 前回の不着事由仕分け２
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【CheckStartUp】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「検査終了」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnStopInspection_Click(object sender, EventArgs e)
        {
            try
            {
                // シリアルデータ送信
                SendSerialData(PubConstClass.CMD_SEND_c);
                LblError.Visible = false;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnStartInspection_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ステータス表示
        /// </summary>
        /// <param name="status"></param>
        private void SetStatus(int status)
        {
            try
            {
                iStatus = status;

                switch (status)
                {
                    case 0:
                        LblStatus.Text = "停止中";
                        LblStatus.BackColor = Color.LightGray;
                        LblStatus.ForeColor = Color.Black;
                        SetControlEnable(true);
                        break;

                    case 1:
                        LblStatus.Text = "検査中";
                        LblStatus.BackColor = Color.LightGreen;
                        LblStatus.ForeColor = Color.Black;
                        SetControlEnable(false);
                        break;

                    case 2:
                        LblStatus.Text = "エラー";
                        LblStatus.BackColor = Color.OrangeRed;
                        LblStatus.ForeColor = Color.White;

                        BtnCounterClear1.Visible = false;
                        BtnCounterClear2.Visible = false;
                        BtnCounterClear3.Visible = false;
                        BtnCounterClear4.Visible = false;
                        BtnCounterClear5.Visible = false;
                        break;

                    case 3:
                        LblStatus.Text = "手動登録中";
                        LblStatus.BackColor = Color.Orange;
                        LblStatus.ForeColor = Color.White;
                        break;

                    default:
                        LblStatus.Text = "停止中";
                        LblStatus.BackColor = Color.LightGray;
                        LblStatus.ForeColor = Color.Black;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【SetStatus】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="bEnable"></param>
        private void SetControlEnable(bool bEnable)
        {
            try
            {
                BtnJobSelect.Enabled = bEnable;
                DtpDateReceipt.Enabled = bEnable;
                CmbNonDeliveryReasonSorting1.Enabled = bEnable;
                CmbNonDeliveryReasonSorting2.Enabled = bEnable;
                BtnSetting.Enabled = bEnable;
                BtnStartInspection.Enabled = bEnable;
                BtnClose.Enabled = bEnable;

                BtnCounterClear1.Visible = bEnable;
                BtnCounterClear2.Visible = bEnable;
                BtnCounterClear3.Visible = bEnable;
                BtnCounterClear4.Visible = bEnable;
                BtnCounterClear5.Visible = bEnable;
                BtnRejectCounterClear.Visible = bEnable;
                BtnAllCounterClear.Visible = bEnable;

                groupBox2.Enabled = bEnable;
                //TxtBoxLabelNumber.Enabled = bEnable;
                //TxtInquiryNumber.Enabled = bEnable;
                //TxtCheckReading.Enabled = bEnable;
                ChkCDCheck.Enabled = bEnable;
                TxtQrReadData.Enabled = bEnable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【SetControlEnable】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// バージョンボタンダブルクリック処理（デバッグ用）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LblVersion_DoubleClick(object sender, EventArgs e)
        {
            if (LblFdrInfo1.Visible == false)
            {
                LblFdrInfo1.Visible = true;
                LblFdrInfo2.Visible = true;
                LblFdrInfo3.Visible = true;
                LblFdrInfo4.Visible = true;
                LblFdrInfo5.Visible = true;
                LblGrpInfo1.Visible = true;
                LblGrpInfo2.Visible = true;
                LblGrpInfo3.Visible = true;
                LblGrpInfo4.Visible = true;
                LblGrpInfo5.Visible = true;
                CmbDigit.Visible = true;
                CmbFontSize.Visible = true;
                TxtTestCounter.Visible = true;
                BtnTestCounter.Visible = true;

                CmbDigit.Items.Clear();
                for (int iIndex = 31; iIndex <= 128; iIndex++)
                {
                    CmbDigit.Items.Add(iIndex.ToString());
                }
                CmbDigit.SelectedIndex = 0;

                CmbFontSize.Items.Clear();
                CmbFontSize.Items.Add("8");
                CmbFontSize.Items.Add("9");
                CmbFontSize.Items.Add("10");
                CmbFontSize.Items.Add("11");
                CmbFontSize.Items.Add("12");
                CmbFontSize.Items.Add("14");
                CmbFontSize.Items.Add("16");
                CmbFontSize.Items.Add("18");
                CmbFontSize.Items.Add("20");
                CmbFontSize.Items.Add("22");
                CmbFontSize.Items.Add("24");
                CmbFontSize.SelectedIndex = 0;
            }
            else
            {
                LblFdrInfo1.Visible = false;
                LblFdrInfo2.Visible = false;
                LblFdrInfo3.Visible = false;
                LblFdrInfo4.Visible = false;
                LblFdrInfo5.Visible = false;
                LblGrpInfo1.Visible = false;
                LblGrpInfo2.Visible = false;
                LblGrpInfo3.Visible = false;
                LblGrpInfo4.Visible = false;
                LblGrpInfo5.Visible = false;
                CmbDigit.Visible = false;
                CmbFontSize.Visible = false;
                TxtTestCounter.Visible = false;
                BtnTestCounter.Visible = false;

                LblPocket1.Text = "";
                LblPocket2.Text = "";
                LblPocket3.Text = "";
                LblPocket4.Text = "";
                LblPocket5.Text = "";
                LblPocketEject.Text = "";
            }
        }

        /// <summary>
        /// 「設定」ボタンクリック処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSetting_Click(object sender, EventArgs e)
        {
            try
            {
                CommonModule.OutPutLogFile("検査画面：「設定」ボタンクリック");
                if (LblSelectedFile.Text.Trim() == "")
                {
                    MessageBox.Show("JOBを選択して下さい", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                //PubConstClass.sJobFileNameFromInspectionForm = "";
                SettingForm form = new SettingForm();
                form.ShowDialog(this);

                // ジョブ登録情報及びグループ１～５情報の読取り
                CommonModule.ReadJobEntryListFile(PubConstClass.sJobFileNameFromInspectionForm);
                // 登録ジョブ項目を取得し表示
                GetEntryInfoAndDisplay();
                // 受領日
                DtpDateReceipt.Enabled = bDateOfReceipt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【SetStatus】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 検査装置からのデータ受信処理
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
                CommonModule.OutPutLogFile($"■受信データ：{data.Replace("\r", "<CR>")}");

                // 受信データの先頭１文字を取得
                string sCommandType = data.Substring(0, 1);
                switch (sCommandType)
                {
                    case PubConstClass.CMD_RECIEVE_A:
                        // JOB設定情報の送信
                        MyProcJobInfomation();
                        break;

                    case PubConstClass.CMD_RECIEVE_B:
                        if (bManualEntryFlg)
                        {
                            // 手動登録中は検査開始不可とする
                            MyProcStop();
                            // シリアルデータ送信
                            SendSerialData(PubConstClass.CMD_SEND_c);
                        }
                        // 開始コマンド
                        if (CheckNumberOfDigits())
                        {
                            // 桁数チェックでOKの場合は検査開始
                            MyProcStart();
                        }
                        break;

                    case PubConstClass.CMD_RECIEVE_C:
                        // 停止コマンド
                        MyProcStop();
                        break;

                    case PubConstClass.CMD_RECIEVE_D:
                        // データコマンド
                        // 先頭2文字（D,）を取り除く
                        MyProcData(data.Substring(2, data.Length - 2));
                        break;

                    case PubConstClass.CMD_RECIEVE_E:
                        // エラーコマンド
                        MyProcError(data);
                        break;

                    case PubConstClass.CMD_RECIEVE_F:
                        // エラーリセットコマンド
                        MyProcErrorReset();
                        break;

                    case PubConstClass.CMD_RECIEVE_L:
                        // QR読取り直後データコマンド
                        MyProcQrData(data.Substring(2, data.Length - 2));
                        break;

                    case PubConstClass.CMD_RECIEVE_T:
                        // DIP-SW 情報送信
                        MyProcDipSw();
                        break;

                    default:
                        // 未定義コマンド
                        CommonModule.OutPutLogFile($"未定義コマンドです：{data.Replace("\r", "<CR>")}");
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
        /// JOB設定情報の送信
        /// </summary>
        private void MyProcJobInfomation()
        {
            string sData;
            
            try
            {
                if (LblSelectedFile.Text.Trim() == "")
                {
                    // JOBが未選択
                    // シリアルデータ送信
                    SendSerialData(PubConstClass.CMD_SEND_e);
                }
                else
                {
                    // フィーダー設定情報
                    string[] sArrayJob = PubConstClass.lstJobEntryList[0].Split(',');
                    //                    0      1             2   3  4   5      6 7 8      9 0 1            2  3 4      5   6  7  8  9  0  1 2 3  
                    // チューリッヒ１ハガキ,ハガキ,2025年1月10日,OFF,47,OFF,物件ID,1,5,届出日,6,8,ファイル区分,14,1,管理No.,15,10,ON,ON,ON,ON,1,1,
                    sData = PubConstClass.CMD_SEND_a + ",";
                    sData += sArrayJob[1] == "ハガキ" ? "0" : "1";    // (01) 媒体           ：1桁
                    sData += ",";
                    sData += sArrayJob[4].PadLeft(3, '0');            // (02) QR桁数         ：2桁→3桁
                    // ラベルのフォントサイズを変更する
                    ChangeLabelFontSize(sArrayJob[4]);
                    sData += ","; 
                    sData += sArrayJob[5] == "OFF" ? "0" : "1";       // (03) 読取チェック   ：1桁
                    sData += ","; 
                    sData += sArrayJob[7].PadLeft(3, '0');            // (04) QR読取項目1開始：2桁→3桁
                    sData += ",";
                    // (05) QR読取項目1桁数
                    //sData += sArrayJob[8].PadLeft(2, '0');            // (05) QR読取項目1桁数：2桁                    
                    if (TxtCheckReading.Text.Trim().Length == 5)
                    {
                        sData += "05";                               // 5桁
                    }
                    else if(TxtCheckReading.Text.Trim().Length == 2)
                    {
                        sData += "02";                               // 2桁
                    }
                    else
                    {
                        sData += "05";                               // 5桁
                    }                    
                    sData += ","; 
                    sData += sArrayJob[10].PadLeft(3, '0');           // (06) QR読取項目2開始：2桁→3桁
                    sData += ","; 
                    sData += sArrayJob[11].PadLeft(2, '0');           // (07) QR読取項目2桁数：2桁                    
                    sData += ","; 
                    sData += sArrayJob[13].PadLeft(3, '0');           // (08) QR読取項目3開始：2桁→3桁
                    sData += ","; 
                    sData += sArrayJob[14].PadLeft(2, '0');           // (09) QR読取項目3桁数：2桁                    
                    sData += ","; 
                    sData += sArrayJob[16].PadLeft(3, '0');           // (10) QR読取項目4開始：2桁→3桁
                    sData += ","; 
                    sData += sArrayJob[17].PadLeft(2, '0');           // (11) QR読取項目4桁数：2桁
                    sData += ",";
                    sData += sArrayJob[18] == "OFF" ? "0" : "1";      // (12) 重複検査　　　 ：1桁
                    sData += ",";
                    sData += sArrayJob[19] == "OFF" ? "0" : "1";      // (13) Ｗフィード検査 ：1桁
                    sData += ",";
                    sData += sArrayJob[20] == "OFF" ? "0" : "1";      // (14) 超音波検知　　 ：1桁
                    sData += ",";
                    sData += sArrayJob[21] == "OFF" ? "0" : "1";      // (15) 桁数チェック　 ：1桁
                    sData += ",";
                    sData += sArrayJob[22];                           // (16) 読取機能　　　 ：1桁
                    sData += ",";
                    sData += sArrayJob[23].PadLeft(3, '0');           // (17) 読取位置　　　 ：3桁
                    sData += ",";
                    sData += ChkCDCheck.Checked == true ? "1" : "0";    // (18) C/Dチェックの有無：1桁
                    sData += ",";

                    // シリアルデータ送信
                    SendSerialData(sData);
                    // コマンドは連続して送信しない
                    Thread.Sleep(150);
                    // ソーター設定のポケット１～５の情報を送信
                    for (int iIndex = 0; iIndex < 5; iIndex++)
                    {
                        MyprocPocket(iIndex);
                        // コマンドは連続して送信しない
                        Thread.Sleep(150);
                    }                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyProcJobInfomation】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ソーター設定のポケット１～５の情報を送信
        /// </summary>
        /// <param name="iPocketNumber">0～4（ポケット１～５）</param>
        private void MyprocPocket(int iPocketNumber)
        {
            string sData;
            int iIndex;

            try
            {
                // ポケット設定情報
                string[] sArrayJob = PubConstClass.lstPocketInfo[0].Split(',');
                //              0 1              2 3        4 5                  6 7                  8 9  0  1  2  3  4  5  6  7  8  9
                // コメリ１ハガキ,1,コメリ２ハガキ,2,武蔵野BK,3,西日本シティーBK１,4,西日本シティーBK２,5,50,50,50,50,50,ON,ON,ON,ON,ON,

                iIndex = int.Parse(sArrayJob[iPocketNumber * 2 + 1]) - 1;

                string[] sArrayPocket;
                string sPocketInfo = "";
                if (iIndex == 5)
                {
                    sArrayPocket = ",,,,,,,,,".Split(',');
                    // イジェクト設定
                    sPocketInfo = "E";
                }
                else
                {
                    sArrayPocket = PubConstClass.lstGroupInfo[iIndex].Split(',');
                    // グループ設定
                    sPocketInfo = (iIndex + 1).ToString("0");
                }

                sData = PubConstClass.CMD_SEND_f + ",";
                sData += (iPocketNumber + 1).ToString("0");                         // ポケット番号
                sData += ",";
                sData += sPocketInfo;                                               // ポケット情報
                sData += ",";
                // ポケット１かどうかのチェック
                if (iPocketNumber == 0)
                {
                    // ポケット１の場合は、物件IDとして「読取値チェック」をセットする
                    sData += TxtCheckReading.Text.Trim();
                }
                else
                {
                    // 物件ID
                    sData += sArrayPocket[1];
                }
                
                sData += ",";
                sData += sArrayPocket[2];                                           // 届出日
                sData += ",";
                sData += sArrayPocket[3];                                           // ファイル区分
                sData += ",";
                sData += sArrayPocket[4];                                           // 管理番号
                sData += ",";
                sData += int.Parse(sArrayJob[10 + iPocketNumber]).ToString("000");  // ポケット切替件数
                sData += ",";
                sData += sArrayJob[15 + iPocketNumber] == "OFF" ? "0": "1";         // ポケット切替ON/OFF
                // シリアルデータ送信
                SendSerialData(sData);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyprocPocket】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 開始コマンド処理
        /// </summary>
        private void MyProcStart()
        {
            try
            {
                if (LblSelectedFile.Text.Trim() == "")
                {
                    // JOBが未選択
                    // シリアルデータ送信
                    SendSerialData(PubConstClass.CMD_SEND_e);
                    return;
                }
                // 検査中
                SetStatus(1);

                TxtBoxLabelNumber.Enabled = false;
                TxtInquiryNumber.Enabled = false;
                TxtCheckReading.Enabled = false;

                // 検査開始時のチェック
                CheckStartUp();
                // JOB設定情報の送信
                MyProcJobInfomation();

                string sPocketCount1 = GetPocketCounter(LblBox1);
                string sPocketCount2 = GetPocketCounter(LblBox2);
                string sPocketCount3 = GetPocketCounter(LblBox3);
                string sPocketCount4 = GetPocketCounter(LblBox4);
                string sPocketCount5 = GetPocketCounter(LblBox5);

                string sData = PubConstClass.CMD_SEND_l + ",";
                sData += sPocketCount1 + ",";
                sData += sPocketCount2 + ",";
                sData += sPocketCount3 + ",";
                sData += sPocketCount4 + ",";
                sData += sPocketCount5 + ",";
                SendSerialData(sData);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyProcStart】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ポケットの表示桁数を取得
        /// </summary>
        /// <param name="label">ポケットの表示桁数</param>
        /// <returns></returns>
        private string GetPocketCounter(Label label)
        {
            string sPocketCount;
            try
            {
                sPocketCount = int.Parse(label.Text).ToString("0000");
                if (label.Text.Length > 4)
                {
                    sPocketCount = "9999";
                }
                return sPocketCount;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【GetPocketCounter】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "0000";
            }
        }

        /// <summary>
        /// 停止コマンド処理
        /// </summary>
        private void MyProcStop()
        {
            try
            {
                // 停止中
                SetStatus(0);
                if (CmbMode.SelectedIndex == 0)
                {
                    // 受付モードの時は検査終了で下記の入力領域は不活性化とする
                    TxtBoxLabelNumber.Enabled = false;
                    TxtInquiryNumber.Enabled = false;
                    TxtCheckReading.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyProcStop】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// エラーコマンド処理
        /// </summary>
        /// <param name="sData"></param>
        private void MyProcError(string sData)
        {
            string sErrorCode;
            string sSaveFileName = "";
            string sErrorData;

            try
            {
                LblError.Text = $"エラーコマンド「{sData.Replace("\r", "<CR>")}」受信";
                LblError.Visible = true;

                sErrorCode = sData.Substring(2, 3);

                if (sErrorCode == "005" || sErrorCode == "013" || sErrorCode == "050")
                {
                    // 停止中（005：用紙終了／013：セットカウントエラー／050：リジェクト停止）
                    SetStatus(0);
                    PubConstClass.bIsErrorMessage = false;
                    LblError.Visible = false;
                }
                else
                {
                    // エラー
                    SetStatus(2);
                    PubConstClass.bIsErrorMessage = true;
                }

                ErrorMessageForm form = ErrorMessageForm.GetInstance();

                sErrorData = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + ",";
                sErrorData += sErrorCode + ",";
                if (PubConstClass.dicErrorCodeData.ContainsKey(sErrorCode))
                {
                    // 存在する場合
                    form.SetMessage($"{sErrorCode},{PubConstClass.dicErrorCodeData[sErrorCode]}");
                    CommonModule.OutPutLogFile($"エラー内容：{sErrorCode},{PubConstClass.dicErrorCodeData[sErrorCode]}");
                    sErrorData += PubConstClass.dicErrorCodeData[sErrorCode];
                }
                else
                {
                    form.SetMessage($"{sErrorCode},未定義エラー番号,未定義のエラー番号です。");
                    CommonModule.OutPutLogFile($"エラー内容：{sErrorCode},未定義エラー番号,未定義のエラー番号です。");
                    sErrorData += "未定義エラー番号,未定義のエラー番号です。";
                }

                //// エラーフォルダ及びエラーファイル名のチェック
                //if (sFolderNameForErrorLog == null || sFileNameForErrorLog == null)
                //{
                //    // NULLの場合
                //    sJobFolderName = CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder);
                //    sJobFolderName += $"エラーログ\\{sProcessingDate}\\";
                //    //sJobFolderName = $"C:\\QRソーター\\エラーログ\\20250816";
                //    sFileNameForErrorLog = $"処理開始前_errorlog_{sReceiptDate}_{DateTime.Now.ToString("yyyyMMdd")}.csv" + "000000.csv";
                //    CommonModule.OutPutLogFile($"エラーファイル名を作成しました：{sFileNameForErrorLog}");
                //    //string sFolderName = "";
                //    //sFolderName += CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder);

                //    //string sFolderName += sFolderNameForErrorLog + sJobFolderName + "\\";
                //    if (!Directory.Exists(sFolderNameForErrorLog))
                //    {
                //        Directory.CreateDirectory(sFolderNameForErrorLog);
                //        CommonModule.OutPutLogFile($"エラーフォルダを作成しました：{sFolderNameForErrorLog}");
                //    }
                //}

                // エラーファイル名の生成
                //sSaveFileName += CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder);
                //sSaveFileName += sFolderNameForErrorLog + sJobFolderName + "\\";
                //sSaveFileName = $"{sFolderNameForErrorLog}\\{sFileNameForErrorLog}";
                sSaveFileName = $"{sFolderNameForErrorLog}\\{sFileNameForErrorLog}";

                // エラーデータ書込処理
                using (StreamWriter sw = new StreamWriter(sSaveFileName, true, Encoding.Default))
                {                    
                    // エラーデータを追加モードで書き込む
                    sw.WriteLine(sErrorData);
                }

                if (!PubConstClass.bIsOpenErrorMessage)
                {
                    PubConstClass.bIsOpenErrorMessage = true;
                    // エラーダイアログ表示
                    form.ShowDialog();
                }
                else
                {
                    form.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyProcError】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// データコマンドの処理
        /// </summary>
        /// <param name="sData"></param>
        private void MyProcData(string sData)
        {
            string[] col = new string[12];
            ListViewItem itm1;
            ListViewItem itm2;
            string[] strArray;
            string sLogData = "";
            string sWriteDate;
            string sWriteTime;
            string sSaveFileName;
            string DQ = "\"";

            try
            {
                if (iStatus == 0)
                {
                    SendSerialData(PubConstClass.CMD_SEND_e);
                    return;
                }
                if (sFolderNameForOkLog == null || sFileNameForOkLog == null)
                {
                    //MessageBox.Show("検査前の設定を行ってください", "確認", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    CommonModule.OutPutLogFile($"データ（{sData}）を受信したが、sFolderNameForOkLog または sFileNameForOkLog が、NULLです");
                    return;
                }

                if (sFolderNameForOkLog == "" || sFileNameForOkLog == "")
                {
                    // 受領日
                    sReceiptDate = DtpDateReceipt.Value.ToString("yyyyMMdd");

                    // 処理日の取得
                    sProcessingDate = DateTime.Now.ToString("yyyyMMdd");

                    // ＯＫ用の検査ログ保存用フォルダの作成
                    sFolderNameForOkLog = CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder) +
                                          sProcessingModeName + "\\" + sProcessingDate;
                    if (Directory.Exists(sFolderNameForOkLog) == false)
                    {
                        Directory.CreateDirectory(sFolderNameForOkLog);
                    }

                    // 全件用の検査ログ保存用フォルダの作成
                    sFolderNameForAllLog = CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder) +
                                           "受付・箱詰め用\\" + sProcessingDate;
                    if (Directory.Exists(sFolderNameForAllLog) == false)
                    {
                        Directory.CreateDirectory(sFolderNameForAllLog);
                    }
                    string sOutPutDateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                    sFileNameForOkLog = $"{sFolderNameForOkLog}\\uketuke_{PubConstClass.pblMachineName}_{sReceiptDate}_{sOutPutDateTime}.csv";
                    sFileNameForAllLog = $"{sFolderNameForAllLog}\\uketuke_{PubConstClass.pblMachineName}_{sReceiptDate}_{sOutPutDateTime}（全件）.csv";

                    CommonModule.OutPutLogFile($"ファイル名作成（{sFileNameForOkLog}／{sFileNameForAllLog}）");

                }

                sWriteDate = DateTime.Now.ToString("yyyy/MM/dd");
                sWriteTime = DateTime.Now.ToString("HH:mm:ss");
               
                strArray = sData.Split(',');
                // 日時
                col[1] = sWriteDate + " " + sWriteTime;
                // 読取値（QRコード）
                //col[2] = strArray[0].Trim();
                col[2] = strArray[0];
                // 判定（OK/NG）
                col[3] = strArray[1].Trim() == "0" ? "OK" : "NG";
                
                // トレイ情報
                col[4] = strArray[3].Trim();

                string sFolderName = "";
                string sFileName = "";
                bool bIsTrayOk = true;

                sFolderName = sFolderNameForOkLog;
                sFileName = sFileNameForOkLog;

                switch (col[4])
                {
                    // トレイ情報の確認
                    case "1":
                        // ポケット１
                        iBox1Count++;
                        LblBox1.Text = iBox1Count.ToString("0");
                        LblPocket1.Text = col[2];
                        break;
                    case "2":
                        // ポケット２
                        iBox2Count++;
                        LblBox2.Text = iBox2Count.ToString("0");
                        LblPocket2.Text = col[2];
                        break;
                    case "3":
                        // ポケット３
                        iBox3Count++;
                        LblBox3.Text = iBox3Count.ToString("0");
                        LblPocket3.Text = col[2];
                        break;
                    case "4":
                        // ポケット４
                        iBox4Count++;
                        LblBox4.Text = iBox4Count.ToString("0");
                        LblPocket4.Text = col[2];
                        break;
                    case "5":
                        // ポケット５
                        iBox5Count++;
                        LblBox5.Text = iBox5Count.ToString("0");
                        LblPocket5.Text = col[2];
                        break;
                    case "E":
                        // リジェクト
                        iBoxECount++;
                        LblBoxEject.Text = iBoxECount.ToString("0");
                        LblPocketEject.Text = col[2];
                        col[3]= "REJECT";
                        break;

                    default:
                        // ポケット情報が不明（イジェクトする）
                        iBoxECount++;
                        LblBoxEject.Text = iBoxECount.ToString("0");
                        LblPocketEject.Text = col[2];
                        col[3] = "ﾎﾟｹｯﾄ不明";
                        bIsTrayOk = false;
                        break;
                }

                // 判定がOK以外でエラー番号（033：重複）なら判定を「重複」とする
                if (strArray[1].Trim() != "0" && strArray[2] == "033")
                {
                    col[3] = "重複";
                }

                sLogData += DQ + sWriteDate.Replace("/", "") + DQ + ",";            // 日付
                sLogData += DQ + sWriteTime + DQ + ",";                             // 時刻
                sLogData += DQ + DQ + ",";                                          // 期待値                       Null
                //sLogData += DQ + strArray[0].Trim() + DQ + ",";                     // 読取値
                sLogData += DQ + strArray[0] + DQ + ",";                            // 読取値
                sLogData += DQ + col[3] + DQ + ",";                                 // 判定
                sLogData += DQ + sFileName + DQ + ",";                              // 正解データファイル名
                sLogData += DQ + DQ + ",";                                          // 重量期待値[g]				Null
                sLogData += DQ + DQ + ",";                                          // 重量測定値[g]				Null
                sLogData += DQ + DQ + ",";                                          // 重量公差						Null
                sLogData += DQ + DQ + ",";                                          // フラップ最大長[mm]			Null
                sLogData += DQ + DQ + ",";                                          // フラップ積算長[mm]			Null
                sLogData += DQ + DQ + ",";                                          // フラップ検出回数[回]			Null
                sLogData += DQ + DQ + ",";                                          // イベント（コメント）			Null
                sLogData += DQ + sReceiptDate + DQ + ",";                           // 受領日
                sLogData += DQ + PubConstClass.sUserId + DQ + ",";                  // 作業者情報                
                if (strArray[0].Length >= intPropertyIdNumber)
                {
                    // 物件IDの切り出し
                    sLogData += DQ + strArray[0].Substring(0, intPropertyIdNumber) + DQ + ",";
                    //CommonModule.OutPutLogFile($"ログ出力時の物件ID = {strArray[0].Substring(0, intPropertyIdNumber)}");
                }
                else
                {
                    // 読取値を物件IDとする
                    sLogData += DQ + strArray[0] + DQ + ",";
                }                
                sLogData += DQ + strArray[2] + DQ + ",";                            // エラー
                sLogData += DQ + DQ + ",";                                          // 生産管理番号					Null
                sLogData += DQ + sBoxLabelNumber + DQ + ",";                        // 箱ラベル番号
                sLogData += DQ + sInquiryNumber + DQ + ",";                         // 問い合わせ番号
                sLogData += DQ + DQ + ",";                                          // ファイル名（画像）			Null
                sLogData += DQ + DQ + ",";                                          // ファイルパス（画像）			Null
                sLogData += DQ + DQ;                                                // 工場コード					Null

                // データの表示（判定が「OK」でトレイ情報が「E」以外）
                if (col[3] == "OK" && col[4] != "E" && bIsTrayOk == true)
                {
                    intOkSesanCounter += 1;
                    // No.
                    col[0] = intOkSesanCounter.ToString("00000");

                    // 重複チェックの検査対象にする
                    lstPastReceivedQrData.Add(col[2]);
                    // OKのカウント表示
                    iOKCount++;
                    LblOKCount.Text = iOKCount.ToString("#,##0");

                    itm1 = new ListViewItem(col);
                    LsvOKHistory.Items.Add(itm1);
                    LsvOKHistory.Items[LsvOKHistory.Items.Count - 1].UseItemStyleForSubItems = false;
                    LsvOKHistory.Select();
                    LsvOKHistory.Items[LsvOKHistory.Items.Count - 1].EnsureVisible();

                    if (LsvOKHistory.Items.Count % 2 == 1)
                    {
                        for (int iIndex = 0; iIndex < 5; iIndex++)
                        {
                            // 奇数行の色反転
                            LsvOKHistory.Items[LsvOKHistory.Items.Count - 1].SubItems[iIndex].BackColor = Color.FromArgb(200, 200, 230);
                        }
                    }

                    if (!Directory.Exists(sFolderNameForOkLog)) 
                    {
                        Directory.CreateDirectory(sFolderNameForOkLog);
                        CommonModule.OutPutLogFile($"OKフォルダを作成しました：{sFolderName}");
                    }

                    // ヘッダー情報書込処理
                    if (!File.Exists(sFileNameForOkLog))
                    {
                        using (StreamWriter sw = new StreamWriter(sFileNameForOkLog, true, Encoding.Default))
                        {
                            // 書込ファイルが無かったらヘッダー情報を書込
                            sw.WriteLine(GetHederInfo());
                        }
                    }
                    // 検査データ書込処理
                    using (StreamWriter sw = new StreamWriter(sFileNameForOkLog, true, Encoding.Default))
                    {
                        // OKデータのみを追加モードで書き込む
                        sw.WriteLine(sLogData);
                    }
                }
                else
                {
                    intNgSesanCounter += 1;
                    // No.
                    col[0] = intNgSesanCounter.ToString("00000");

                    // NGのカウント表示
                    iNGCount++;
                    LblNGCount.Text = iNGCount.ToString("#,##0");

                    itm2 = new ListViewItem(col);
                    LsvNGHistory.Items.Add(itm2);
                    LsvNGHistory.Items[LsvNGHistory.Items.Count - 1].UseItemStyleForSubItems = false;
                    LsvNGHistory.Select();
                    LsvNGHistory.Items[LsvNGHistory.Items.Count - 1].EnsureVisible();
                    if (LsvNGHistory.Items.Count % 2 == 1)
                    {
                        for (int iIndex = 0; iIndex < 5; iIndex++)

                        {
                            // 奇数行の色反転
                            LsvNGHistory.Items[LsvNGHistory.Items.Count - 1].SubItems[iIndex].BackColor = Color.FromArgb(200, 200, 230);
                        }
                    }
                }

                // 総数のカウント表示
                LblTotalCount.Text = (iOKCount + iNGCount).ToString("#,##0");

                // 900セット以上の時は、50で割り切れるかをチェックする
                if (iOKCount >= 900)
                {
                    if (iOKCount % 50 == 0)
                    {
                        LblOffLine.BackColor = Color.WhiteSmoke;
                        // 900、950、1000、1050、110、、と50単位で停止する。
                        // シリアルデータ送信
                        SendSerialData(PubConstClass.CMD_SEND_c);
                        LblError.Visible = false;

                        MyProcStop();
                    }
                }

                // ヘッダー情報書込処理
                if (!File.Exists(sFileNameForAllLog))
                {
                    using (StreamWriter sw = new StreamWriter(sFileNameForAllLog, true, Encoding.Default))
                    {
                        // 書込ファイルが無かったらヘッダー情報を書込
                        sw.WriteLine(GetHederInfo());
                    }
                }
                // 検査データ書込処理
                using (StreamWriter sw = new StreamWriter(sFileNameForAllLog, true, Encoding.Default))
                {
                    // 全件データを追加モードで書き込む
                    sw.WriteLine(sLogData);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyProcData】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CommonModule.OutPutLogFile("【MyProcData】" + ex.Message);
            }
        }

        /// <summary>
        /// QR読取り直後データの表示
        /// </summary>
        /// <param name="sData"></param>
        private void MyProcQrData(string sData)
        {
            string sQrData;
            try
            {
                //LblQrReadData.Text = sData.Replace("\r","<CR>");
                //sQrData = sData.Replace("\r", "").Trim();
                sQrData = sData.Replace("\r", "");
                LblQrReadData.Text = sQrData;

                if (lstPastReceivedQrData.Count > 0)
                {
                    #region 重複チェック
                    if (bIsDuplicateCheck)
                    {
                        Stopwatch sw = new Stopwatch();
                        sw.Start();
                        bool bFind = lstPastReceivedQrData.Contains(sQrData);
                        sw.Stop();

                        if (bFind)
                        {
                            CommonModule.OutPutLogFile($"重複データ：{sQrData}");
                            // シリアルデータ送信（重複エラー発生）
                            SendSerialData(PubConstClass.CMD_SEND_g1);
                        }
                        else
                        {
                            CommonModule.OutPutLogFile($"重複データ無し：{lstPastReceivedQrData[0]}");
                            // シリアルデータ送信（重複エラー発生）
                            SendSerialData(PubConstClass.CMD_SEND_g0);
                        }
                        CommonModule.OutPutLogFile($"{lstPastReceivedQrData.Count:#,###,##0}件の検索処理時間: {sw.Elapsed.TotalMilliseconds}ミリ秒");
                    }
                    else
                    {
                        CommonModule.OutPutLogFile($"重複チェック無しモード：{sQrData}");
                        // シリアルデータ送信（重複エラー発生）
                        SendSerialData(PubConstClass.CMD_SEND_g0);
                    }
                    #endregion
                }
                else
                {
                    CommonModule.OutPutLogFile($"最初のデータなので重複無し：{sQrData}");
                    // シリアルデータ送信（重複エラー無しモード）
                    SendSerialData(PubConstClass.CMD_SEND_g0);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyProcQrData】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CommonModule.OutPutLogFile("【MyProcQrData】" + ex.Message);
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
        /// エラーリセットコマンド送信
        /// </summary>
        private void MyProcErrorReset()
        {
            try
            {
                ErrorMessageForm form = ErrorMessageForm.GetInstance();

                if (form.Visible == true)
                {
                    // エラー表示をクリア
                    LblError.Visible = false;
                    // 停止中
                    SetStatus(0);
                    // エラーメッセージウィンドウを隠す
                    form.Hide();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【MyProcErrorReset】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CommonModule.OutPutLogFile("【MyProcErrorReset】" + ex.Message);
            }
        }

        /// <summary>
        /// ヘッダーデータの取得
        /// </summary>
        /// <returns></returns>
        private string GetHederInfo()
        {
            string sHeader = "";
            
            try
            {                
                sHeader += "\"日付\",";
                sHeader += "\"時刻\",";
                sHeader += "\"期待値\",";
                sHeader += "\"読取値\",";
                sHeader += "\"判定\",";
                sHeader += "\"正解データファイル名\",";
                sHeader += "\"重量期待値[g]\",";
                sHeader += "\"重量測定値[g]\",";
                sHeader += "\"重量公差\",";
                sHeader += "\"フラップ最大長[mm]\",";
                sHeader += "\"フラップ積算長[mm]\",";
                sHeader += "\"フラップ検出回数[回]\",";
                sHeader += "\"イベント（コメント）\",";
                sHeader += "\"受領日\",";
                sHeader += "\"作業員情報（機械情報）\",";
                //sHeader += "\"物件情報（DPS/BPO/Broad等）\",";
                sHeader += "\"市区町村コード\",";
                sHeader += "\"エラーコード\",";
                sHeader += "\"生産管理番号\",";
                //sHeader += "\"仕分けコード１\",";
                //sHeader += "\"仕分けコード２\",";
                sHeader += "\"箱ラベル番号\",";
                sHeader += "\"問い合わせ番号\",";
                sHeader += "\"ファイル名（画像）\",";
                sHeader += "\"ファイルパス（画像）\",";
                sHeader += "\"工場コード\"";
                return sHeader;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【GetHederInfo】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return sHeader;
            }
        }

        /// <summary>
        /// シリアルデータ受信イベント（漢字データ受信対応）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SerialPortQr_DataReceived_Kanji(object sender, SerialDataReceivedEventArgs e)
        {
            string data = "";

            try
            {
                // シリアルポートをオープンしていない場合、処理を行わない。
                if (SerialPortQr.IsOpen == false)
                    return;

                SerialPort sp = (SerialPort)sender;
                while (sp.BytesToRead > 0)
                {
                    byte b = (byte)sp.ReadByte();
                    buffer[bufferIndex++] = b;
                    // CRまで読み込む
                    if (b == '\r')
                    {
                        data = Encoding.GetEncoding("Shift_JIS").GetString(buffer, 0, bufferIndex);
                        Console.WriteLine("Data Received: " + data);
                        bufferIndex = 0;
                        BeginInvoke(new Delegate_RcvDataToTextBox(RcvDataToTextBox), data + "\r");
                    }
                }
            }
            catch (TimeoutException)
            {
                // ディスカードするデータ
                CommonModule.OutPutLogFile("データ受信タイムアウトエラー：<CR>未受信で切り捨てたデータ：" + data);
            }
            catch (Exception ex)
            {
                CommonModule.OutPutLogFile("【SerialPortBcr_DataReceived】" + ex.Message);
            }
        }

        /// <summary>
        /// 固定のジョブファイル（国勢調査用JOB設定.csv）の読込処理
        /// </summary>
        private void LoadingFixedJobFile()
        {
            try
            {
                // 固定の国勢調査用JOB設定ファイルを読み込み 
                string sSelectedFile = CommonModule.IncludeTrailingPathDelimiter(Application.StartupPath) + @"国勢調査用JOB設定.csv";

                string[] sArray = sSelectedFile.Split('\\');
                // ファイル名のみを表示する
                LblSelectedFile.Text = sArray[sArray.Length - 1];

                // ジョブ登録情報及びグループ１～５情報の読取り
                CommonModule.ReadJobEntryListFile(sSelectedFile);
                // 登録ジョブ項目を取得し表示
                GetEntryInfoAndDisplay();
                // 受領日
                DtpDateReceipt.Enabled = bDateOfReceipt;
                // 「検査開始」ボタン使用可
                BtnStartInspection.Enabled = true;
                // 「設定」ボタン使用可
                BtnSetting.Enabled = true;
                PubConstClass.sJobFileNameFromInspectionForm = sSelectedFile;
                // JOB変更フラグON
                bIsJobChange = true;
                // 各表示カウンタクリア
                LblTotalCount.Text = "0";
                LblOKCount.Text = "0";
                LblNGCount.Text = "0";
                // ポケット１～５の表示カウンタクリア
                LblBox1.Text = "0";
                LblBox2.Text = "0";
                LblBox3.Text = "0";
                LblBox4.Text = "0";
                LblBox5.Text = "0";
                LblBoxEject.Text = "0";
                // 内部カウンタのクリア
                iOKCount = 0;               // OK用カウンタ
                iNGCount = 0;               // NG用カウンタ
                iBox1Count = 0;             // ボックス１用カウンタ
                iBox2Count = 0;             // ボックス２用カウンタ
                iBox3Count = 0;             // ボックス３用カウンタ
                iBox4Count = 0;             // ボックス４用カウンタ
                iBox5Count = 0;             // ボックス５用カウンタ
                iBoxECount = 0;             // ボックス（Eject）用カウンタ
                intOkSesanCounter = 0;      // OK処理数No.カウンタ
                intNgSesanCounter = 0;      // NG処理数No.カウンタ
                                            // 受信データ表示領域のクリア
                LblPocket1.Text = "";
                LblPocket2.Text = "";
                LblPocket3.Text = "";
                LblPocket4.Text = "";
                LblPocket5.Text = "";
                LblPocketEject.Text = "";
                // OK履歴とNG履歴のクリア
                LsvOKHistory.Items.Clear();
                LsvNGHistory.Items.Clear();

                // 過去に受信したQRデータ一覧のクリア
                lstPastReceivedQrData.Clear();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【LoadingFixedJobFile】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        
        }

        /// <summary>
        /// 「JOB選択」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnJobSelect_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();

                CommonModule.OutPutLogFile("「JOB選択」ボタンクリック");
                // 初期表示するフォルダの指定（「空の文字列」の時は現在のディレクトリを表示）
                //ofd.InitialDirectory = @"C:\";
                // 「ファイルの種類」に表示される選択肢の指定
                ofd.Filter = "CSVファイル(*.csv;*.CSV)|*.csv;*.CSV|すべてのファイル(*.*)|*.*";
                // 「ファイルの種類」ではじめに「CSVファイル(*.csv;*.CSV)」を選択
                ofd.FilterIndex = 1;
                // タイトルを設定
                ofd.Title = "JOB設定ファイルを選択してください";
                // ダイアログボックスを閉じる前に現在のディレクトリを復元
                ofd.RestoreDirectory = true;
                // 存在しないファイルの名前が指定されたとき警告を表示
                ofd.CheckFileExists = true;
                // 存在しないパスが指定されたとき警告を表示
                ofd.CheckPathExists = true;
                // ダイアログを表示する
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    // 「OK」ボタンがクリック（選択されたファイル名を表示）
                    string sSelectedFile = ofd.FileName;
                    string[] sArray = sSelectedFile.Split('\\');
                    // ファイル名のみを表示する
                    LblSelectedFile.Text = sArray[sArray.Length - 1];
                    // ジョブ登録情報及びグループ１～５情報の読取り
                    CommonModule.ReadJobEntryListFile(sSelectedFile);
                    // 登録ジョブ項目を取得し表示
                    GetEntryInfoAndDisplay();
                    // 受領日
                    DtpDateReceipt.Enabled = bDateOfReceipt;
                    // 「検査開始」ボタン使用可
                    BtnStartInspection.Enabled = true;
                    // 「設定」ボタン使用可
                    BtnSetting.Enabled = true;
                    PubConstClass.sJobFileNameFromInspectionForm = sSelectedFile;

                    // 内部カウンタと表示をクリアする
                    ClearCounterAndDisplay();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnJobSelect_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 内部カウンタと表示をクリアする
        /// </summary>
        private void ClearCounterAndDisplay()
        {
            try
            {

                sProcessingModeName = "";     // 処理モード名
                sProcessingDate = "";         // 処理日
                sBoxLabelNumber = "";         // 箱ラベル番号
                sInquiryNumber = "";          // 問い合わせ番号 
                sReceiptDate = "";            // 受領日
                sFolderNameForOkLog = "";     // OK用の操作ログファイル名
                sFolderNameForAllLog = "";    // 全件用の操作ログファイル名
                //sFolderNameForErrorLog = "";  // エラーログファイル名
                sFileNameForOkLog = "";       // OK用の操作ログファイル名
                sFileNameForAllLog = "";      // 全件用の操作ログファイル名
                //sFileNameForErrorLog = "";    // エラーログファイル名

                TxtBoxLabelNumber.Text = "";
                TxtInquiryNumber.Text = "";
                TxtCheckReading.Text = "";
                TxtQrReadData.Text = "";
                LblQrReadData.Text = "";

                // JOB変更フラグON
                bIsJobChange = true;
                // 各表示カウンタクリア
                LblTotalCount.Text = "0";
                LblOKCount.Text = "0";
                LblNGCount.Text = "0";
                // ポケット１～５の表示カウンタクリア
                LblBox1.Text = "0";
                LblBox2.Text = "0";
                LblBox3.Text = "0";
                LblBox4.Text = "0";
                LblBox5.Text = "0";
                LblBoxEject.Text = "0";
                // 内部カウンタのクリア
                iOKCount = 0;               // OK用カウンタ
                iNGCount = 0;               // NG用カウンタ
                iBox1Count = 0;             // ボックス１用カウンタ
                iBox2Count = 0;             // ボックス２用カウンタ
                iBox3Count = 0;             // ボックス３用カウンタ
                iBox4Count = 0;             // ボックス４用カウンタ
                iBox5Count = 0;             // ボックス５用カウンタ
                iBoxECount = 0;             // ボックス（Eject）用カウンタ
                intOkSesanCounter = 0;      // OK処理数No.カウンタ
                intNgSesanCounter = 0;      // NG処理数No.カウンタ
                                            // 受信データ表示領域のクリア
                LblPocket1.Text = "";
                LblPocket2.Text = "";
                LblPocket3.Text = "";
                LblPocket4.Text = "";
                LblPocket5.Text = "";
                LblPocketEject.Text = "";
                // OK履歴とNG履歴のクリア
                LsvOKHistory.Items.Clear();
                LsvNGHistory.Items.Clear();

                // 過去に受信したQRデータ一覧のクリア
                lstPastReceivedQrData.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【ClearCounterAndDisplay】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 物件ID桁数（デフォルト：5桁）
        private int intPropertyIdNumber = 5;
        /// <summary>
        /// 
        /// </summary>
        private void GetEntryInfoAndDisplay()
        {
            string[] sArray;
            try
            {
                sArray = PubConstClass.lstJobEntryList[0].Split(',');
                // 物件ID桁数
                intPropertyIdNumber = int.Parse(sArray[8]);
                CommonModule.OutPutLogFile($"物件ID桁数(intPropertyIdNumber) = {intPropertyIdNumber}");

                // 受領日
                DtpDateReceipt.Text = sArray[2];
                // 受領日入力
                bDateOfReceipt = sArray[3] == "ON";
                // 重複チェック
                if (sArray[18] == "ON")
                {
                    LblDuplicateCheck.Text = "重複チェック：有";
                    bIsDuplicateCheck = true;
                }
                else
                {
                    LblDuplicateCheck.Text = "重複チェック：無";
                    bIsDuplicateCheck = false;
                }

                LstSettingInfomation.Items.Clear();
                LstSettingInfomation.Items.Add("【設定内容】");
                LstSettingInfomation.Items.Add($"Ｗフィード検査：{sArray[19]}");
                //LstSettingInfomation.Items.Add($"超音波検査　　：{sArray[20]}");
                LstSettingInfomation.Items.Add($"桁数チェック　：{sArray[21]}");
                LstSettingInfomation.Items.Add($"読取機能　　　：{PubConstClass.lstReadFunctionList[int.Parse(sArray[22])]}");
                LstSettingInfomation.Items.Add($"読取チェック　：{sArray[5]}");
                LstSettingInfomation.Items.Add($"読取位置　　　：{sArray[23]} mm");
                //LstSettingInfomation.Items.Add("C/D チェック　：");

                sArray = PubConstClass.lstPocketInfo[0].Split(',');
                // ポケット①名称
                LblBoxTitle1.Text = "【BOX1】 " + sArray[0];
                // ポケット②名称
                LblBoxTitle2.Text = "【BOX2】 " + sArray[2];
                // ポケット③名称
                LblBoxTitle3.Text = "【BOX3】 " + sArray[4];
                // ポケット④名称
                LblBoxTitle4.Text = "【BOX4】 " + sArray[6];
                // ポケット⑤名称
                LblBoxTitle5.Text = "【BOX5】 " + sArray[8];

                // ポケット１切替件数
                LblQuantity1.Text = sArray[15] == "ON" ? sArray[10] : "---";
                // ポケット２切替件数
                LblQuantity2.Text = sArray[16] == "ON" ? sArray[11] : "---";
                // ポケット３切替件数
                LblQuantity3.Text = sArray[17] == "ON" ? sArray[12] : "---";
                // ポケット４切替件数
                LblQuantity4.Text = sArray[18] == "ON" ? sArray[13] : "---";
                // ポケット５切替件数
                LblQuantity5.Text = sArray[19] == "ON" ? sArray[14] : "---";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【GetEntryInfoAndDisplay】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    

        private static bool bIsReset = false;
        public static void SendResetCommand()
        {
            try
            {
                bIsReset = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【SendResetCommand】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ポケットのカウンタクリア
        /// </summary>
        /// <param name="sMessage"></param>
        /// <param name="label"></param>
        /// <param name="iBoxCounter"></param>
        private void CounterClear(string sMessage, Label label, ref int iBoxCounter)
        {
            try
            {
                DialogResult result = MessageBox.Show($"{sMessage}のカウンタをクリアしますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    label.Text = "0";
                    iBoxCounter = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【CounterClear】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCounterClear1_Click(object sender, EventArgs e)
        {
            CounterClear("ポケット１", LblBox1, ref iBox1Count);
        }

        private void BtnCounterClear2_Click(object sender, EventArgs e)
        {
            CounterClear("ポケット２", LblBox2, ref iBox2Count);
        }

        private void BtnCounterClear3_Click(object sender, EventArgs e)
        {
            CounterClear("ポケット３", LblBox3, ref iBox3Count);
        }

        private void BtnCounterClear4_Click(object sender, EventArgs e)
        {
            CounterClear("ポケット４", LblBox4, ref iBox4Count);
        }

        private void BtnCounterClear5_Click(object sender, EventArgs e)
        {
            CounterClear("ポケット５", LblBox5, ref iBox5Count);
        }
        private void BtnRejectCounterClear_Click(object sender, EventArgs e)
        {
            CounterClear("リジェクト", LblBoxEject, ref iBoxECount);
        }

        private void CmbDigit_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sData = "AB3456789*CD3456789*EF3456789*GH3456789*IJ3456789*KL3456789*MN3456789*OP3456789*QR3456789*ST3456789*UV3456789*WX3456789*YZ345678";
            try
            {
                LblQrReadData.Text = sData.Substring(0, int.Parse(CmbDigit.Text));
                LblPocket1.Text = sData.Substring(0, int.Parse(CmbDigit.Text));
                LblPocket2.Text = sData.Substring(0, int.Parse(CmbDigit.Text));
                LblPocket3.Text = sData.Substring(0, int.Parse(CmbDigit.Text));
                LblPocket4.Text = sData.Substring(0, int.Parse(CmbDigit.Text));
                LblPocket5.Text = sData.Substring(0, int.Parse(CmbDigit.Text));
                LblPocketEject.Text = sData.Substring(0, int.Parse(CmbDigit.Text));               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【CmbDigit_SelectedIndexChanged】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CmbFontSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ChangeFontSize(LblQrReadData, float.Parse(CmbFontSize.Text));
                ChangeFontSize(LblPocket1, float.Parse(CmbFontSize.Text));
                ChangeFontSize(LblPocket2, float.Parse(CmbFontSize.Text));
                ChangeFontSize(LblPocket3, float.Parse(CmbFontSize.Text));
                ChangeFontSize(LblPocket4, float.Parse(CmbFontSize.Text));
                ChangeFontSize(LblPocket5, float.Parse(CmbFontSize.Text));
                ChangeFontSize(LblPocketEject, float.Parse(CmbFontSize.Text));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【CmbFontSize_SelectedIndexChanged】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// ラベルのフォントサイズの変更
        /// </summary>
        /// <param name="label"></param>
        /// <param name="newSize"></param>
        private void ChangeFontSize(Label label, float newSize)
        {
            try 
            {
                // 元のフォント情報を取得して、新しいサイズで再作成
                label.Font = new Font(label.Font.FontFamily, newSize, label.Font.Style);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【ChangeFontSize】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// ラベルのフォントサイズの変更
        /// </summary>
        /// <param name="sDigit"></param>
        private void ChangeLabelFontSize(string sDigit)
        {
            int iDigit;
            float fFontSize = 0;

            try
            {
                // 桁数でフォントの大きさを変更
                iDigit = int.Parse(sDigit);
                
                if (iDigit <= 52)
                {
                    // 1桁数～52桁
                    fFontSize = 14.0f;                    
                }
                else if(iDigit <= 60)
                {
                    // 53桁数～60桁
                    fFontSize = 12.0f;
                }
                else if (iDigit <= 66)
                {
                    // 61桁数～66桁
                    fFontSize = 10.0f;
                }
                else if (iDigit <= 128)
                {
                    // 67桁数～128桁
                    fFontSize = 9.0f;
                }
                
                CommonModule.OutPutLogFile($"フォントサイズ：{fFontSize}");
                ChangeFontSize(LblPocket1, fFontSize);
                ChangeFontSize(LblPocket2, fFontSize);
                ChangeFontSize(LblPocket3, fFontSize);
                ChangeFontSize(LblPocket4, fFontSize);
                ChangeFontSize(LblPocket5, fFontSize);
                ChangeFontSize(LblPocketEject, fFontSize);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【ChangeLabelFontSize】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnTestCounter_Click(object sender, EventArgs e)
        {
            LblBox1.Text = TxtTestCounter.Text;
            LblBox2.Text = TxtTestCounter.Text;
            LblBox3.Text = TxtTestCounter.Text;
            LblBox4.Text = TxtTestCounter.Text;
            LblBox5.Text = TxtTestCounter.Text;
        }

        private void BtnAllCounterClear_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = MessageBox.Show("ポケット１～５とリジェクトのカウンタをクリアしますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    LblBox1.Text = "0";
                    LblBox2.Text = "0";
                    LblBox3.Text = "0";
                    LblBox4.Text = "0";
                    LblBox5.Text = "0";
                    LblBoxEject.Text = "0";
                    iBox1Count = 0;
                    iBox2Count = 0;
                    iBox3Count = 0;
                    iBox4Count = 0;
                    iBox5Count = 0;
                    iBoxECount = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnAllCounterClear_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int iPreviousIndex = 1;

        private void CmbMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            //string sMessage = "国勢調査用アプリモード" + Environment.NewLine;
            string sMessage = "";

            try
            {
                // 背景色をリセットする
                LblOffLine.BackColor = Color.WhiteSmoke;

                if (iPreviousIndex == CmbMode.SelectedIndex)
                {
                    // 前回値と同じなら何もせずに抜ける
                    return;
                }

                DialogResult dialogResult = MessageBox.Show($"モードを「{CmbMode.Text}」に切り替えますか？","【モード切り替え確認】", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Cancel)
                {
                    CmbMode.SelectedIndex = iPreviousIndex;
                    return;
                }

                iPreviousIndex = CmbMode.SelectedIndex;

                // 内部カウンタと表示をクリアする
                ClearCounterAndDisplay();

                if (CmbMode.SelectedIndex == 0)
                {
                    sMessage = "受付モード";
                    TxtBoxLabelNumber.Enabled = false;
                    TxtInquiryNumber.Enabled = false;
                    TxtCheckReading.Enabled = false;
                    TxtQrReadData.Enabled = false;

                    // 受付モード
                    sProcessingModeName = "受付用";

                    // 受領日
                    sReceiptDate = DtpDateReceipt.Value.ToString("yyyyMMdd");

                    // 処理日の取得
                    sProcessingDate = DateTime.Now.ToString("yyyyMMdd");

                    // ＯＫ用の検査ログ保存用フォルダの作成
                    sFolderNameForOkLog = CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder) +
                                          sProcessingModeName + "\\" + sProcessingDate;
                    if (Directory.Exists(sFolderNameForOkLog) == false)
                    {
                        Directory.CreateDirectory(sFolderNameForOkLog);
                    }

                    // 全件用の検査ログ保存用フォルダの作成
                    sFolderNameForAllLog = CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder) +
                                           "受付・箱詰め用\\" + sProcessingDate;
                    if (Directory.Exists(sFolderNameForAllLog) == false)
                    {
                        Directory.CreateDirectory(sFolderNameForAllLog);
                    }

                    string sOutPutDateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                    sFileNameForOkLog = $"{sFolderNameForOkLog}\\uketuke_{PubConstClass.pblMachineName}_{sReceiptDate}_{sOutPutDateTime}.csv";
                    sFileNameForAllLog = $"{sFolderNameForAllLog}\\uketuke_{PubConstClass.pblMachineName}_{sReceiptDate}_{sOutPutDateTime}（全件）.csv";
                }
                else
                {
                    sMessage = "箱詰めモード";
                    TxtBoxLabelNumber.Enabled = true;
                    TxtInquiryNumber.Enabled = true;
                    TxtCheckReading.Enabled = true;
                    TxtQrReadData.Enabled = false;
                }
                LblOffLine.Text = sMessage;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【CmbMode_SelectedIndexChanged】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TxtBoxLabelNumber_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    SetTxtCheckReading();

                    //string input = TxtBoxLabelNumber.Text;
                    //string firstFive = input.Length >= 5 ? input.Substring(0, 5) : input;
                    //TxtCheckReading.Text = firstFive;

                    //e.SuppressKeyPress = true; // Enterキーの「ピンッ」という音を防ぐ
                    //TxtInquiryNumber.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【TxtBoxLabelNumber_KeyDown】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TxtQrReadData_Key(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    //e.SuppressKeyPress = true; // Enterキーの「ピンッ」という音を防ぐ

                    if (CmbMode.SelectedIndex == 1)
                    {
                        // 箱詰めモードの場合
                        // 各入力フィールドの桁数チェックを行う
                        if (CheckNumberOfDigits() == false)
                        {
                            //// シリアルデータ送信
                            //SendSerialData(PubConstClass.CMD_SEND_b);
                            //// 検査開始時のチェック
                            //CheckStartUp();
                            return;
                        }
                    }

                    string input = TxtQrReadData.Text;
                    if (!(input.Length == 16 || input.Length == 20))
                    {
                        MessageBox.Show($"入力したQRデータ（{input}) は、16桁あるいは、20桁のデータを読み取って下さい", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        TxtQrReadData.Text = "";
                        TxtQrReadData.Focus();
                        return;
                    }

                    if (lstPastReceivedQrData.Count > 0)
                    {
                        #region 重複チェック
                        if (bIsDuplicateCheck)
                        {
                            Stopwatch sw = new Stopwatch();
                            sw.Start();
                            bool bFind = lstPastReceivedQrData.Contains(input);
                            sw.Stop();

                            if (bFind)
                            {
                                CommonModule.OutPutLogFile($"重複データ：{input}");
                                MessageBox.Show($"入力したQRデータ（{input}) は重複しています", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                TxtQrReadData.Text = "";
                                TxtQrReadData.Focus();
                                return;
                            }
                            else
                            {
                                CommonModule.OutPutLogFile($"重複データ無し：{lstPastReceivedQrData[0]}");
                            }
                            CommonModule.OutPutLogFile($"{lstPastReceivedQrData.Count:#,###,##0}件の検索処理時間: {sw.Elapsed.TotalMilliseconds}ミリ秒");
                        }
                        else
                        {
                            CommonModule.OutPutLogFile($"重複チェック無しモード：{input}");
                        }
                        #endregion
                    }
                    else
                    {
                        CommonModule.OutPutLogFile($"最初のデータなので重複無し：{input}");
                    }

                    if (ChkCDCheck.Checked == true)
                    {
                        string sCheckDigit = GetCheckDigit(input);
                        // if (sCheckDigit != input.Substring(input.Length - 1, 1))
                        if (sCheckDigit != input.Substring(15, 1))
                        {
                            MessageBox.Show($"チェックデジット（{sCheckDigit}）が異なります", "確認", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            TxtQrReadData.Text = "";
                            TxtQrReadData.Focus();
                            return;
                        }
                    }

                    if (TxtCheckReading.Text.Trim() != "")
                    {
                        // 照合チェック
                        int iKeta = TxtCheckReading.Text.Trim().Length;
                        if (TxtQrReadData.Text.Substring(0, iKeta) != TxtCheckReading.Text)
                        {
                            MessageBox.Show($"照合エラー：（{TxtCheckReading.Text}）が異なります", "確認", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            TxtQrReadData.Text = "";
                            TxtQrReadData.Focus();
                            return;
                        }
                    }

                    // 手動登録モードON
                    bManualEntryFlg = true;


                    DialogResult dialogResult = MessageBox.Show($"読取りデータ（{input}）を登録しますか？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Cancel)
                    {
                        TxtQrReadData.Text = "";
                        TxtQrReadData.Focus();
                        // 手動登録モードOFF
                        bManualEntryFlg = false;
                        return;
                    }
                    else
                    {
                        TxtQrReadData.Text = "";
                        TxtQrReadData.Focus();
                        // 手動登録モードOFF
                        bManualEntryFlg = false;
                        bIsJobChange = false;
                    }
                    //// 手動登録モードOFF
                    //bManualEntryFlg = false;

                    // 登録処理
                    string sData = $"{input},0,000,1,\r";
                    //iStatus = 3;
                    SetStatus(3);
                    MyProcData(sData);
                    TxtQrReadData.Text = "";
                    TxtQrReadData.Focus();
                    //iStatus = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【TxtQrReadData_Key】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetCheckDigit(string input)
        {
            string[] aryCheckDigit = new string[36] {"0","1","2","3","4","5","6","7","8","9",
                                                     "A","B","C","D","E","F","G","H","I","J",
                                                     "K","L","M","N","O","P","Q","R","S","T",
                                                     "U","V","W","X","Y","Z"};
            try
            {
                if (input.Length < 16)
                {
                    return input;
                }
                string s1 = input.Substring(0, 5);
                string s2 = input.Substring(5, 4);
                string s3 = input.Substring(9, 1);
                string s4 = input.Substring(10, 2);
                string s5 = input.Substring(12, 3);

                int iTotal = int.Parse(s1) + int.Parse(s2) + int.Parse(s3) + int.Parse(s4) + int.Parse(s5);
                int iMod = iTotal % 36;

                return aryCheckDigit[iMod];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【TxtQrReadData_Key】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "ERROR!!";
            }
        }

        private void BtnJobChange_Click(object sender, EventArgs e)
        {
            try
            {
                // 背景色をリセットする
                LblOffLine.BackColor = Color.WhiteSmoke;

                DialogResult dialogResult = MessageBox.Show("JOBを切り替えますか？","確認",MessageBoxButtons.OKCancel,MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Cancel)
                {
                    return;
                }

                TxtBoxLabelNumber.Enabled = true;
                TxtInquiryNumber.Enabled = true;
                TxtCheckReading.Enabled = true;

                string sOutPutDateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                sFileNameForOkLog = $"{sFolderNameForOkLog}\\uketuke_{PubConstClass.pblMachineName}_{sReceiptDate}_{sOutPutDateTime}.csv";
                sFileNameForAllLog = $"{sFolderNameForAllLog}\\uketuke_{PubConstClass.pblMachineName}_{sReceiptDate}_{sOutPutDateTime}（全件）.csv";
                // 内部カウンタと表示をクリアする
                ClearCounterAndDisplay();

                if (CmbMode.SelectedIndex == 0)
                {
                    // 受付モード
                    sProcessingModeName = "受付用";
                }
                else
                {
                    // 箱詰めモード
                    sProcessingModeName = "箱詰め用";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnJobChange_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TxtInquiryNumber_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true; // Enterキーの「ピンッ」という音を防ぐ
                    //TxtCheckReading.Focus();
                    BtnStartInspection.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【TxtInquiryNumber_KeyDown】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TxtCheckReading_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true; // Enterキーの「ピンッ」という音を防ぐ
                    TxtBoxLabelNumber.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【TxtCheckReading_KeyDown】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TxtBoxLabelNumber_Leave(object sender, EventArgs e)
        {
            SetTxtCheckReading();
        }

        private void SetTxtCheckReading()
        {
            try
            {
                string input = TxtBoxLabelNumber.Text;
                string firstFive = input.Length >= 5 ? input.Substring(0, 5) : input;
                TxtCheckReading.Text = firstFive;

                //e.SuppressKeyPress = true; // Enterキーの「ピンッ」という音を防ぐ
                TxtInquiryNumber.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【SetTxtCheckReading】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
