using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QrSorterInspectionApp
{
    internal class PubConstClass
    {
        /// <summary>
        /// アプリから制御CPUへのコマンド
        /// </summary>
        public const string CMD_SEND_a = "a";                                   // JOBの設定内容送信コマンド
        public const string CMD_SEND_b = "b";                                   // 検査開始コマンド
        public const string CMD_SEND_c = "c";                                   // 検査終了コマンド
        public const string CMD_SEND_d = "d";                                   // エラーリセットコマンド
        public const string CMD_SEND_e = "e";                                   // 動作不可コマンド
        public const string CMD_SEND_f = "f";                                   // ポケット投入条件情報コマンド        
        public const string CMD_SEND_g = "g";                                   // 重複エラー発生コマンド(g0,g1分類する)
        public const string CMD_SEND_g0 = "g0";                                 // 重複エラー無しコマンド
        public const string CMD_SEND_g1 = "g1";                                 // 重複エラー有りコマンド
        public const string CMD_SEND_h = "h";                                   // JOB選択コマンド
        public const string CMD_SEND_i = "i";                                   // アワーメータークリアコマンド
        public const string CMD_SEND_j = "j";                                   // トータルカウンタクリアコマンド
        public const string CMD_SEND_k = "k";                                   // 保守画面の出力チェックコマンド
        public const string CMD_SEND_l = "l";                                   // 検査画面カウンタ送信コマンド
        public const string CMD_SEND_m = "m";                                   // m,0：通常モード／m,1：メンテナンスモード
        public const string CMD_SEND_t = "t";                                   // パネルDIP-SW情報送信コマンド
        /// <summary>
        /// 制御CPUからアプリへのコマンド
        /// </summary>
        public const string CMD_RECIEVE_A = "A";                                // 検査状況確認コマンド
        public const string CMD_RECIEVE_B = "B";                                // 検査開始コマンド
        public const string CMD_RECIEVE_C = "C";                                // 検査終了コマンド
        public const string CMD_RECIEVE_D = "D";                                // ポケット投入時QR読取りデータ
        public const string CMD_RECIEVE_E = "E";                                // エラー番号送信コマンド
        public const string CMD_RECIEVE_F = "F";                                // エラーリセット送信コマンド
        public const string CMD_RECIEVE_I = "I";                                // アワーメーター送信コマンド
        public const string CMD_RECIEVE_J = "J";                                // トータルカウンタ送信コマンド
        public const string CMD_RECIEVE_K = "K";                                // I/O状態送信コマンド
        public const string CMD_RECIEVE_L = "L";                                // QR読取り直後のデータ送信コマンド
        public const string CMD_RECIEVE_T = "T";                                // DIP-SW情報要求コマンド

        public const string DEF_VERSION = "Ver.0.8.1.2";                        // バージョン情報（メジャー.マイナー.ビルド.リビジョン） 
        public static object objSyncHist;                                       // 排他制御用オブジェクト（操作ログ書込用）

        /// <summary>
        /// ログイン画面
        /// </summary>
        public static string sUserId;                                           // ログイン中のユーザーID
        public static string sUserName;                                         // ログイン中のユーザー名
        public static string sUserPassword;                                     // ログイン中のパスワード
        public static string sUserAuthority;                                    // ログイン中のユーザー権限

        /// <summary>
        /// メニュー画面
        /// </summary>
        public const string DEF_ERROR_FILE = "ErrorMessage.txt";                // エラーメッセージファイル名称
        public static Dictionary<string, string> dicErrorCodeData;              // エラーコード変換用辞書
        public const string DEF_READ_FUNCTION__FILE = "ReadFunctionItem.txt";   // 読取機能項目ファイル名称
        public static List<string> lstReadFunctionList = new List<string>();    // 読取機能項目情報

        /// <summary>
        /// 検査画面
        /// </summary>
        public static string sPrevDtpDateReceipt = "";                          // 前回の受領日
        public static string sPrevNonDelivery1 = "";                            // 前回の不着事由仕分け１
        public static string sPrevNonDelivery2 = "";                            // 前回の不着事由仕分け２
        public static bool   bIsOpenErrorMessage = false;                       // エラーメッセージ画面表示・非表示フラグ
        public static bool   bIsErrorMessage = false;                           // true：エラーメッセージ／false：情報メッセージ
        /// <summary>
        /// 設定画面        
        /// </summary>
        public const string DEF_JOB_ENTRY_FILE_NAME = "JobEntryList.txt";       // 登録ジョブファイル名
        public const string DEF_NON_DELIVERY = "NonDeliveryReasonSorting.txt";  // 不着事由格納ファイル
        public static List<string> lstJobEntryList = new List<string>();        // 登録ジョブ情報
        public static List<string> lstGroupInfo = new List<string>();           // グループ１～５の情報
        public static List<string> lstPocketInfo = new List<string>();          // ポケット１～５の情報
        public static List<string> lstNonDeliveryList = new List<string>();     // 不着事由情報
        public static string sJobFileNameFromInspectionForm = "";               // 検査画面呼出時のジョブファイル名

        /// <summary>
        /// SV・OP画面
        /// </summary>
        public static List<string> lstUserAccount = new List<string>();         // ユーザーアカウント情報
        public const string DEF_FILENAME = "QrSorterInspectionApp.def";         // DEFファイル名称
        public const string DEF_USER_ACCOUNT_FILE_NAME = "UserAccount.txt";     // 復号化ユーザーアカウントファイル名
        public const string DEF_USER_ACCOUNT_ENC_FILE_NAME = "UserAccount.enc"; // 暗号化ユーザーアカウントファイル名
        public const string DEF_DES_KEY = "QRINSPCT";                           // 暗号・復号用キー（8文字）

        /// <summary>
        /// 保守画面
        /// </summary>
        // 保守項目
        public const string DEF_MACHINE_NAME         = "号機名称";
        public const string DEF_LOGSAVE_MONTH        = "ログ保存期間";
        public const string DEF_HDD_SPACE            = "ディスク空き容量";
        public const string DEF_INTERNAL_TRAN_FOLDER = "内部実績ログ格納フォルダ";
        public const string DEF_DIP_SW               = "DIPSW";
        public const string DEF_OFFLINE_MODE         = "オフラインモード";
        // COMポート１
        public const string DEF_COMPORT              = "COMポート名";
        public const string DEF_COM_SPEED            = "COM通信速度";
        public const string DEF_COM_DATA_LENGTH      = "COMデータ長";
        public const string DEF_COM_IS_PARITY        = "COMパリティ有無";
        public const string DEF_COM_PARITY_VAR       = "COMパリティ種別";
        public const string DEF_COM_STOPBIT          = "COMストップビット";
        // 保守項目
        public static string pblMachineName;                                    // 号機名                
        public static string pblSaveLogMonth;                                   // ログ保存期間
        public static string pblHddSpace;                                       // ディスク空き容量
        public static string pblInternalTranFolder;                             // 内部実績ログ格納フォルダ
        public static string pblDipSw;                                          // DIP-SW
        public static string pblOffLineMode;                                    // オフラインモード
        // COMポート１
        public static string pblComPort;                                        // COMポート名
        public static string pblComSpeed;                                       // 通信速度
        public static string pblComDataLength;                                  // データ長（0：8bit／1：7bit）
        public static string pblComIsParity;                                    // パリティの有無（0：無効／1：有効）
        public static string pblComParityVar;                                   // パリティ種別（0：奇数／1：偶数）
        public static string pblComStopBit;                                     // ストップビット（0：1bit／1：2bit）
    }
}
