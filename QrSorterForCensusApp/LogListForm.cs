using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace QrSorterInspectionApp
{
    public partial class LogListForm : Form
    {
        public LogListForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// フォームロード処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LogListForm_Load(object sender, EventArgs e)
        {
            try
            {
                LblVersion.Text = PubConstClass.DEF_VERSION;
                CommonModule.OutPutLogFile("ログ画面を表示しました");
                LblSelectedFile.Text = "";
                CmbLogType.Items.Clear();
                CmbLogType.Items.Add("ＯＫログ");
                CmbLogType.Items.Add("全件ログ");
                CmbLogType.SelectedIndex = 0;

                #region 検査ログ一覧のヘッダー設定
                // ListViewのカラムヘッダー設定
                LsvLogList.View = View.Details;
                ColumnHeader col1 = new ColumnHeader();
                ColumnHeader col2 = new ColumnHeader();
                ColumnHeader col3 = new ColumnHeader();
                col1.Text = "　　検査ログファイル名";
                col2.Text = "件数";
                col3.Text = "格納フォルダ";
                col1.TextAlign = HorizontalAlignment.Left;
                col2.TextAlign = HorizontalAlignment.Center;
                col3.TextAlign = HorizontalAlignment.Left;
                col1.Width = 600;         // 検査ログファイル名
                col2.Width = 100;         // 件数
                col3.Width = 1100;        // 格納フォルダ
                ColumnHeader[] colHeaderList = new[] { col1, col2, col3 };
                LsvLogList.Columns.AddRange(colHeaderList);
                #endregion
                #region 検査ログ内容のヘッダー設定
                // ListViewのカラムヘッダー設定
                LsvLogContent.View = View.Details;
                ColumnHeader col01 = new ColumnHeader();
                ColumnHeader col02 = new ColumnHeader();
                ColumnHeader col03 = new ColumnHeader();
                ColumnHeader col04 = new ColumnHeader();
                ColumnHeader col05 = new ColumnHeader();
                ColumnHeader col06 = new ColumnHeader();
                ColumnHeader col07 = new ColumnHeader();
                ColumnHeader col08 = new ColumnHeader();
                ColumnHeader col09 = new ColumnHeader();
                ColumnHeader col10 = new ColumnHeader();
                ColumnHeader col11 = new ColumnHeader();
                col01.Text = "　　日付";
                col02.Text = "時刻";
                col03.Text = "読取値";
                col04.Text = "判定";
                col05.Text = "正解データファイル名";
                col06.Text = "受領日";
                col07.Text = "作業員情報";
                col08.Text = "物件情報";    // 物件情報（DPS/BPO/Broad等）
                col09.Text = "エラーCD";
                col10.Text = "仕分①";
                col11.Text = "仕分②";
                col01.TextAlign = HorizontalAlignment.Center;
                col02.TextAlign = HorizontalAlignment.Center;
                col03.TextAlign = HorizontalAlignment.Center;
                col04.TextAlign = HorizontalAlignment.Center;
                col05.TextAlign = HorizontalAlignment.Center;
                col06.TextAlign = HorizontalAlignment.Center;
                col07.TextAlign = HorizontalAlignment.Center;
                col08.TextAlign = HorizontalAlignment.Center;
                col09.TextAlign = HorizontalAlignment.Center;
                col10.TextAlign = HorizontalAlignment.Center;
                col11.TextAlign = HorizontalAlignment.Center;
                col01.Width = 120;         // 日付
                col02.Width = 100;         // 時刻
                col03.Width = 340;         // 読取値
                col04.Width = 80;          // 判定
                col05.Width = 560;         // 正解データファイル名
                col06.Width = 110;         // 受領日
                col07.Width = 110;         // 作業員情報（機械情報）
                col08.Width = 110;         // 物件情報（DPS/BPO/Broad等）
                col09.Width = 90;          // エラーコード
                col10.Width = 80;          // 仕分けコード①
                col11.Width = 80;          // 仕分けコード②
                ColumnHeader[] colHeaderOK = new[] { col01, col02, col03, col04, col05,
                                                     col06, col07, col08, col09, col10,
                                                     col11
                                                   };
                LsvLogContent.Columns.AddRange(colHeaderOK);
                #endregion
                #region 不着事由仕分け１と２の項目設定
                CommonModule.ReadNonDeliveryList();
                CmbReasonForNonDelivery1.Items.Clear();
                CmbReasonForNonDelivery1.Items.Clear();
                foreach (var items in PubConstClass.lstNonDeliveryList)
                {
                    CmbReasonForNonDelivery1.Items.Add(items);
                    CmbReasonForNonDelivery2.Items.Add(items);
                }
                CmbReasonForNonDelivery1.SelectedIndex = 0;
                CmbReasonForNonDelivery2.SelectedIndex = 0;
                #endregion

                LblLogFileCount.Text = "";
                LblContentCount.Text = "";
                LblSelectedFile.Text = "";

                CmbSortBy.Items.Clear();
                CmbSortBy.Items.Add("ファイル作成順");
                CmbSortBy.Items.Add("ファイル名順");
                CmbSortBy.SelectedIndex = 0;

                //EnableDoubleBuffering(LsbLogList);
                //EnableDoubleBuffering(LsvLogContent);
                
                // 検査ログ一覧表示処理
                //InspectionLogList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【LogListForm_Load】", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                Owner.Show();
                Owner.Refresh();
                this.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnClose_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ログファイル一覧格納リスト
        private List<string> lstLogFileList = new List<string>();

        /// <summary>
        /// 検査ログデータの１行分の表示
        /// </summary>
        /// <param name="sData"></param>
        private void DisplayOneData(string sData)
        {
            // ●（01）日付
            // ●（02）時刻
            // 　（03）期待値
            // ●（04）読取値
            // ●（05）判定
            // ●（06）正解データファイル名
            // 　（07）重量期待値[g]
            // 　（08）重量測定値[g]
            // 　（09）重量公差
            // 　（10）フラップ最大長[mm]
            // 　（11）フラップ積算長[mm]
            // 　（12）フラップ検出回数[回]
            // 　（13）イベント（コメント）
            // ●（14）受領日
            // ●（15）作業員情報（機械情報）
            // ●（16）物件情報（DPS/BPO/Broad等）
            // ●（17）エラーコード
            // 　（18）生産管理番号
            // ●（19）仕分けコード１
            // ●（20）仕分けコード２
            // 　（21）ファイル名（画像）
            // 　（22）ファイルパス（画像）
            // 　（23）工場コード
            try
            {
                PicWaitContent.Refresh();
                //Application.DoEvents();
                string[] sArray = sData.Split(',');
                // "日付","期待値","読取値","判定","正解データファイル名","重量期待値[g]","重量測定値[g]","重量公差","フラップ最大長[mm]","フラップ積算長[mm]","フラップ検出回数[回]","イベント（コメント）","受領日","作業員情報（機械情報）","物件情報（DPS/BPO/Broad等）","エラーコード","生産管理番号","仕分けコード１","仕分けコード２","ファイル名（画像）","ファイルパス（画像）","工場コード",
                string[] col = new string[11];                                              
                ListViewItem itm;
                col[0] = sArray[0].Substring(1, sArray[0].Length - 2);      // 日付
                col[1] = sArray[1].Substring(1, sArray[1].Length - 2);      // 時刻
                col[2] = sArray[3].Substring(1, sArray[3].Length - 2);      // 読取値
                col[3] = sArray[4].Substring(1, sArray[4].Length - 2);      // 判定
                col[4] = sArray[5].Substring(1, sArray[5].Length - 2);      // 正解データファイル名
                col[5] = sArray[13].Substring(1, sArray[13].Length - 2);    // 受領日
                col[6] = sArray[14].Substring(1, sArray[14].Length - 2);    // 作業員情報
                col[7] = sArray[15].Substring(1, sArray[15].Length - 2);    // 物件情報
                col[8] = sArray[16].Substring(1, sArray[16].Length - 2);    // エラーCD
                col[9] = sArray[18].Substring(1, sArray[18].Length - 2);    // 仕分①
                col[10] = sArray[19].Substring(1, sArray[19].Length - 2);   // 仕分②

                // データの表示
                itm = new ListViewItem(col);
                LsvLogContent.Items.Add(itm);
                LsvLogContent.Items[LsvLogContent.Items.Count - 1].UseItemStyleForSubItems = false;
                LsvLogContent.Select();
                LsvLogContent.Items[LsvLogContent.Items.Count - 1].EnsureVisible();

                if (LsvLogContent.Items.Count % 2 == 1)
                {
                    for (int iIndex = 0; iIndex < 11; iIndex++)
                    {
                        // 奇数行の色反転
                        LsvLogContent.Items[LsvLogContent.Items.Count - 1].SubItems[iIndex].BackColor = Color.FromArgb(200, 200, 230);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【DisplayOneData】", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                    // 検査ログ一覧表示処理
                    //InspectionLogList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnJobSelect_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 検査ログ一覧表示処理
        /// </summary>
        private void InspectionLogList()
        {
            string[] sArray;
            string[] sArrayJob;
            string sPath;
            //string sMes;
            string sTitle;

            try
            {                
                if (CmbLogType.SelectedIndex == 0)
                {
                    sPath = "QRソーター設定検査ログ（OKのみ）\\";
                    //sMes = "（OKのみ）";
                    sTitle = "（OKのみ）検査ログファイル件数：";
                }
                else
                {
                    sPath = "QRソーター設定検査ログ（全件）\\";
                    //sMes = "（全件）";
                    sTitle = "（全件）検査ログファイル件数：";
                }

                if (LblSelectedFile.Text != "")
                {
                    sArrayJob = LblSelectedFile.Text.Split('.');
                    sPath += sArrayJob[0] + "\\";
                }
                else
                {
                    sArrayJob = ".csv".Split('.');
                    sPath += "\\";
                }

                if (!Directory.Exists(CommonModule.IncludeTrailingPathDelimiter(PubConstClass.pblInternalTranFolder) + sPath))
                {                    
                    CommonModule.OutPutLogFile($"【ログ管理】JOB（{sArrayJob[0]}）は、未検査のJOBです");
                    MessageBox.Show($"JOB（{sArrayJob[0]}）は、未検査のJOBです", "確認", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // ログファイル一覧格納リストのクリア
                lstLogFileList.Clear();
                // 検査ログ一覧のクリア
                LsvLogList.Items.Clear();
                // 検査ログの内容のクリア
                LsvLogContent.Items.Clear();

                List<string> lstFileList = new List<string>();
                lstFileList.Clear();
                if (CmbSortBy.SelectedIndex == 0)
                {
                    // ファイル作成順
                    foreach (string sTranFile in Directory.GetFiles(CommonModule.IncludeTrailingPathDelimiter(
                                                                          PubConstClass.pblInternalTranFolder) +
                                                                          sPath,
                                                                          "*", SearchOption.AllDirectories).OrderByDescending(f => File.GetLastWriteTime(f)))
                    {
                        lstFileList.Add(sTranFile);
                    }
                }
                else
                {
                    // ファイル名順
                    foreach (string sTranFile in Directory.GetFiles(CommonModule.IncludeTrailingPathDelimiter(
                                                                          PubConstClass.pblInternalTranFolder) +
                                                                          sPath,
                                                                          "*", SearchOption.AllDirectories))
                    {
                        lstFileList.Add(sTranFile);
                    }
                }

                // 検査ログ対象ファイルの取得
                foreach (string sTranFile in lstFileList)
                {
                    PicWaitList.Refresh();
                    //CommonModule.OutPutLogFile($"{sMes}検査ログ対象ファイル：{sTranFile}");
                    sArray = sTranFile.Split('\\');
                    string sFileName = sArray[sArray.Length - 1];
                    string sFileNameFullPath = sTranFile;
                    // 検査日付で絞り込む
                    if (ChkInspectionDate.Checked)
                    {
                        string[] sArrayDate = sFileName.Split('_');
                        if (!(int.Parse(dtTimePickerFrom.Value.ToString("yyyyMMdd")) <= int.Parse(sArrayDate[sArrayDate.Length-1].Substring(0, 8)) &
                            int.Parse(dtTimePickerTo.Value.ToString("yyyyMMdd")) >= int.Parse(sArrayDate[sArrayDate.Length - 1].Substring(0, 8))))
                        {
                            // 該当しないので対象ファイルから外す
                            sFileName = "";
                            sFileNameFullPath = "";
                        }
                    }
                    // 不着事由仕分け１で絞り込む
                    if (ChkReasonForNonDelivery1.Checked)
                    {
                        if (sFileName != "")
                        {
                            string[] sArrayNonDeli1 = sFileName.Split('_');
                            string[] sArrayCmbNonDeli1 = CmbReasonForNonDelivery1.Text.Split(',');

                            if (int.Parse(sArrayCmbNonDeli1[0]) != int.Parse(sArrayNonDeli1[sArrayNonDeli1.Length-4]))
                            {
                                // 該当しないので対象ファイルから外す
                                sFileName = "";
                                sFileNameFullPath = "";
                            }

                        }
                    }
                    // 不着事由仕分け２で絞り込む
                    if (ChkReasonForNonDelivery2.Checked)
                    {
                        if (sFileName != "")
                        {
                            string[] sArrayNonDeli2 = sFileName.Split('_');
                            string[] sArrayCmbNonDeli2 = CmbReasonForNonDelivery2.Text.Split(',');

                            if (int.Parse(sArrayCmbNonDeli2[0]) != int.Parse(sArrayNonDeli2[sArrayNonDeli2.Length-3]))
                            {
                                // 該当しないので対象ファイルから外す
                                sFileName = "";
                                sFileNameFullPath = "";
                            }
                        }
                    }

                    if (sFileName != "")
                    {
                        string sPathName;


                        if (CmbLogType.SelectedIndex == 0)
                        {
                            // OKログ
                            if (sArray[3] == "")
                            {
                                // JOB名でのフィルタ無し
                                sPathName = sArray[0] + "¥" + sArray[1] + "¥" + sArray[2] + "¥" + sArray[4] + "¥" + sArray[5];
                            }
                            else
                            {
                                // JOB名でのフィルタ有り
                                sPathName = sArray[0] + "¥" + sArray[1] + "¥" + sArray[2] + "¥" + sArray[3] + "¥" + sArray[4];
                            }
                        }
                        else
                        {
                            // 全件ログ                           
                            if (sArray[3] == "")
                            {
                                // JOB名でのフィルタ無し
                                sPathName = sArray[0] + "¥" + sArray[1] + "¥" + sArray[2] + "¥" + sArray[4];
                            }
                            else
                            {
                                // JOB名でのフィルタ有り
                                sPathName = sArray[0] + "¥" + sArray[1] + "¥" + sArray[2] + "¥" + sArray[3];
                            }
                        }
                        // 件数の取得
                        string[] Lines = File.ReadAllLines(sTranFile);
                        // 検査ログファイル一覧格納リストに追加
                        lstLogFileList.Add(sTranFile);

                        string[] col = new string[3];
                        ListViewItem itm;
                        col[0] = sArray[sArray.Length - 1];     // ファイル名
                        col[1] = $"{Lines.Length - 1}件";       // 件数
                        col[2] = sPathName;                     // 格納フォルダ

                        // データの表示
                        itm = new ListViewItem(col);
                        LsvLogList.Items.Add(itm);
                        LsvLogList.Items[0].UseItemStyleForSubItems = false;
                        LsvLogList.Select();
                        LsvLogList.Items[0].EnsureVisible();
                    }
                }

                if (sArrayJob[0] == "")
                {
                    sArrayJob[0] = "指定なし";
                }
                LblLogFileCount.Text = $"JOB名（{sArrayJob[0]}）{sTitle}{LsvLogList.Items.Count:#,###} 件";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【InspectionLogList】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnJobClear_Click(object sender, EventArgs e)
        {
            LblSelectedFile.Text = "";
        }

        /// <summary>
        /// 「更新」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            SetEnableControl(false);
            
            PicWaitList.Visible = true;
            PicWaitList.Refresh();
            // 検査ログ一覧表示処理
            InspectionLogList();
            PicWaitList.Visible = false;
            
            SetEnableControl(true);
        }

        /// <summary>
        /// コントロールの有効／無効設定
        /// </summary>
        /// <param name="bEnabled"></param>
        private void SetEnableControl(bool bEnabled)
        {
            try
            {
                CmbLogType.Enabled = bEnabled;
                BtnJobSelect.Enabled = bEnabled;
                BtnJobClear.Enabled = bEnabled;
                GrpSortBy.Enabled = bEnabled;
                BtnUpdate.Enabled = bEnabled;
                GrpInspectionDate.Enabled = bEnabled;
                GrpReasonForNonDelivery1.Enabled = bEnabled;
                GrpReasonForNonDelivery2.Enabled = bEnabled;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【SetEnableControl】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 検査ログ一覧の選択クリック処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LsvLogList_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sReadLogFile;
            string sData;
            int iCounter;

            try
            {
                if (LsvLogList.SelectedItems.Count == 0)
                {
                    return;
                }

                LsvLogContent.Items.Clear();

                sReadLogFile = lstLogFileList[LsvLogList.SelectedItems[0].Index];

                SetEnableControl(false);
                PicWaitContent.Visible = true;
                iCounter = 0;
                PubConstClass.lstJobEntryList.Clear();
                using (StreamReader sr = new StreamReader(sReadLogFile, Encoding.Default))
                {
                    while (!sr.EndOfStream)
                    {
                        //PicWaitContent.Refresh();
                        sData = sr.ReadLine();
                        if (iCounter > 0)
                        {
                            DisplayOneData(sData);
                        }
                        else
                        {
                            //CommonModule.OutPutLogFile($"ヘッダー情報をスキップ：{sData}");
                            CommonModule.OutPutLogFile("ヘッダー情報をスキップ");
                        }
                        iCounter++;
                    }
                }
                SetEnableControl(true);
                PicWaitContent.Visible = false;
                LblContentCount.Text = $"表示ログ件数：{LsvLogContent.Items.Count:#,###} 件";

                LsvLogContent.Items[0].UseItemStyleForSubItems = false;
                LsvLogContent.Select();
                LsvLogContent.Items[0].EnsureVisible();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【LsvLogList_SelectedIndexChanged】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
