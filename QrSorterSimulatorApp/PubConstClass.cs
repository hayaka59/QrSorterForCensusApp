using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QrSorterSimulatorApp
{
    internal class PubConstClass
    {
        public const string CMD_SEND_A = "A";      // 
        public const string CMD_SEND_B = "B";      // 
        public const string CMD_SEND_C = "C";      // 
        public const string CMD_SEND_D = "D";      // 
        public const string CMD_SEND_E = "E";      // 
        public const string CMD_SEND_F = "F";      // 
        public const string CMD_SEND_I = "I";      // アワーメーター
        public const string CMD_SEND_J = "J";      // トータルカウンタ
        public const string CMD_SEND_K = "K";      // 
        public const string CMD_SEND_L = "L";      // 
        public const string CMD_SEND_T = "T";      // DIP-DW情報

        public const string CMD_RECIEVE_a = "a";   // 
        public const string CMD_RECIEVE_b = "b";   // 
        public const string CMD_RECIEVE_c = "c";   // 
        public const string CMD_RECIEVE_d = "d";   // 
        public const string CMD_RECIEVE_e = "e";   // 

        public const string DEF_VERSION = "Ver.0.4.0.5";                        // バージョン情報（メジャー.マイナー.ビルド.リビジョン） 
        public static object objSyncHist;

        public const string DEF_FILENAME = "QrSorterSimulatorApp.def";          // DEFファイル名称
        public static List<string> lstNonDeliveryList = new List<string>();     // 不着事由情報
        public const string DEF_NON_DELIVERY = "NonDeliveryReasonSorting.txt";  // 不着事由格納ファイル

        // 保守画面
        // COMポート
        public const string DEF_COMPORT =         "COMポート名";
        public const string DEF_COM_SPEED =       "COM通信速度";
        public const string DEF_COM_DATA_LENGTH = "COMデータ長";
        public const string DEF_COM_IS_PARITY =   "COMパリティ有無";
        public const string DEF_COM_PARITY_VAR =  "COMパリティ種別";
        public const string DEF_COM_STOPBIT =     "COMストップビット";

        public static string pblMainFormTitle;
        // COMポート
        public static string pblComPort;                                        // COMポート名
        public static string pblComSpeed;                                       // 通信速度
        public static string pblComDataLength;                                  // データ長（0：8bit／1：7bit）
        public static string pblComIsParity;                                    // パリティの有無（0：無効／1：有効）
        public static string pblComParityVar;                                   // パリティ種別（0：奇数／1：偶数）
        public static string pblComStopBit;                                     // ストップビット（0：1bit／1：2bit）
    }
}
