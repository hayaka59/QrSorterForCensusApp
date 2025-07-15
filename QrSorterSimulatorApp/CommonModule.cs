using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QrSorterSimulatorApp
{
    internal class CommonModule
    {
        /// <summary>
        /// フォルダの末尾の「\」を保証する
        /// </summary>
        /// <param name="strCheckFolder">チェック対象のフォルダ名称</param>
        /// <returns>チェック後のフォルダ名称</returns>
        /// <remarks></remarks>
        public static string IncludeTrailingPathDelimiter(string strCheckFolder)
        {
            string IncludeTrailingPathDelimiter;
            try
            {
                if (strCheckFolder.Substring(strCheckFolder.Length - 1, 1) == @"\")
                    IncludeTrailingPathDelimiter = strCheckFolder;
                else
                    IncludeTrailingPathDelimiter = strCheckFolder + @"\";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【IncludeTrailingPathDelimiter】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "";
            }
            return IncludeTrailingPathDelimiter;
        }

        /// <summary>
        /// 操作履歴ログの書込処理
        /// </summary>
        /// <param name="strOutPutData">操作履歴メッセージ</param>
        /// <remarks></remarks>
        public static void OutPutLogFile(string strOutPutData)
        {
            string strOutPutFolder;
            string strOutPutFileName;
            string strYYYYMMDDvalue;
            string strHHMMSSvalue;
            string strPutMessage;

            lock ((PubConstClass.objSyncHist))
            {
                try
                {
                    DateTime dtNow = DateTime.Now;

                    // 指定した書式で日付を文字列に変換する
                    string strNowFormat = dtNow.ToString("yyyy/MM/dd HH:mm:ss");

                    {
                        var withBlock = DateTime.Now;
                        strYYYYMMDDvalue = string.Format("{0:D4}{1:D2}{2:D2}", withBlock.Year, withBlock.Month, withBlock.Day);
                        strHHMMSSvalue = string.Format("{0:D2}{1:D2}{2:D2}", withBlock.Hour, withBlock.Minute, withBlock.Second);
                    }

                    // 格納フォルダ名の設定
                    strOutPutFolder = IncludeTrailingPathDelimiter(Application.StartupPath) + @"OPHISTORYLOG\";
                    // 格納ファイル名の設定
                    strOutPutFileName = "操作履歴ログ_" + strYYYYMMDDvalue + ".LOG";

                    if (Directory.Exists(strOutPutFolder) == false)
                    {
                        // フォルダの作成
                        Directory.CreateDirectory(strOutPutFolder);
                        strPutMessage = strNowFormat + "【" + strOutPutFolder + "】フォルダを作成しました。";
                        // 追記モードで書き込む
                        using (StreamWriter sw = new StreamWriter(strOutPutFolder + strOutPutFileName, true, Encoding.Default))
                        {
                            sw.WriteLine(strPutMessage);
                        }
                    }

                    // 操作履歴ログに操作ログ内容を書き込む
                    strPutMessage = strNowFormat + "：" + strOutPutData;
                    // 追記モードで書き込む
                    using (StreamWriter sw = new StreamWriter(strOutPutFolder + strOutPutFileName, true, Encoding.Default))
                    {
                        sw.WriteLine(strPutMessage);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "【OutPutLogFile】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


        /// <summary>
        /// システム定義ファイルの読込処理
        /// </summary>
        /// <remarks></remarks>
        public static void ReadSystemDefinition()
        {
            string strReadDataPath;
            string[] strArray;

            try
            {
                strReadDataPath = IncludeTrailingPathDelimiter(Application.StartupPath) + PubConstClass.DEF_FILENAME;

                using (StreamReader sr = new StreamReader(strReadDataPath, Encoding.Default))
                {
                    while (!sr.EndOfStream)
                    {
                        strArray = sr.ReadLine().Split(',');
                        switch (strArray[0])
                        {
                            // COMポート名
                            case PubConstClass.DEF_COMPORT:
                                {
                                    PubConstClass.pblComPort = strArray[1];
                                    break;
                                }
                            // COM通信速度
                            case PubConstClass.DEF_COM_SPEED:
                                {
                                    PubConstClass.pblComSpeed = strArray[1];
                                    break;
                                }
                            // COMデータ長
                            case PubConstClass.DEF_COM_DATA_LENGTH:
                                {
                                    PubConstClass.pblComDataLength = strArray[1];
                                    break;
                                }
                            // COMパリティ有無
                            case PubConstClass.DEF_COM_IS_PARITY:
                                {
                                    PubConstClass.pblComIsParity = strArray[1];
                                    break;
                                }
                            // COMパリティ種別
                            case PubConstClass.DEF_COM_PARITY_VAR:
                                {
                                    PubConstClass.pblComParityVar = strArray[1];
                                    break;
                                }
                            // COMストップビット
                            case PubConstClass.DEF_COM_STOPBIT:
                                {
                                    PubConstClass.pblComStopBit = strArray[1];
                                    break;
                                }
                            default:
                                {
                                    break;
                                }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【ReadSystemDefinition】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// システム定義ファイルの書込処理
        /// </summary>
        /// <remarks></remarks>
        public static void WriteSystemDefinition()
        {
            string strPutDataPath;

            try
            {
                strPutDataPath = IncludeTrailingPathDelimiter(Application.StartupPath) + PubConstClass.DEF_FILENAME;

                // 上書モードで書き込む
                using (StreamWriter sw = new StreamWriter(strPutDataPath, false, Encoding.Default))
                {
                    // COMポート名
                    sw.WriteLine(PubConstClass.DEF_COMPORT + "," + PubConstClass.pblComPort);
                    // COM通信速度
                    sw.WriteLine(PubConstClass.DEF_COM_SPEED + "," + PubConstClass.pblComSpeed);
                    // COMデータ長
                    sw.WriteLine(PubConstClass.DEF_COM_DATA_LENGTH + "," + PubConstClass.pblComDataLength);
                    // COMパリティ有無
                    sw.WriteLine(PubConstClass.DEF_COM_IS_PARITY + "," + PubConstClass.pblComIsParity);
                    // COMパリティ種別
                    sw.WriteLine(PubConstClass.DEF_COM_PARITY_VAR + "," + PubConstClass.pblComParityVar);
                    // COMストップビット
                    sw.WriteLine(PubConstClass.DEF_COM_STOPBIT + "," + PubConstClass.pblComStopBit);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【WritetSystemDefinition】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 不着事由情報ファイル読込
        /// </summary>
        public static void ReadNonDeliveryList()
        {
            string sReadDataPath;
            string sData;

            try
            {
                sReadDataPath = IncludeTrailingPathDelimiter(Application.StartupPath) + PubConstClass.DEF_NON_DELIVERY;

                PubConstClass.lstNonDeliveryList.Clear();
                using (StreamReader sr = new StreamReader(sReadDataPath, Encoding.Default))
                {
                    while (!sr.EndOfStream)
                    {
                        sData = sr.ReadLine();
                        PubConstClass.lstNonDeliveryList.Add(sData);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【ReadNonDeliveryList】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
