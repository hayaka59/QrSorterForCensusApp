using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace QrSorterInspectionApp
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// フォームロード処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LoginForm_Load(object sender, EventArgs e)
        {
            try
            {
                // バージョン表示
                LblVersion.Text = PubConstClass.DEF_VERSION;
                PubConstClass.objSyncHist = new object();
                CommonModule.OutPutLogFile("〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓");
                CommonModule.OutPutLogFile("【" + "QRソータ検査アプリ" + "】を起動しました。");
                CommonModule.OutPutLogFile("QRソータ検査アプリバージョン「" + PubConstClass.DEF_VERSION + "」");

                // 設定ファイルの存在チェック
                if (ConfigurationFileExistenceCheck() == false)
                {
                    // 設定ファイルがない場合
                    this.Dispose();
                }
                // 二重起動のチェック
                if (System.Diagnostics.Process.GetProcessesByName(System.Diagnostics.Process.GetCurrentProcess().ProcessName).Length > 1)
                {
                    // すでに起動していると判断する
                    CommonModule.OutPutLogFile("二重起動はできません。");
                    MessageBox.Show("二重起動はできません。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    this.Dispose();
                }

                // 暗号化されたユーザーアカウントファイルの読込
                CommonModule.ReadEncodeUserAccountFile();
                // IMEモードの禁止
                TxtUserId.ImeMode = ImeMode.Disable;
                TxtPassword.ImeMode = ImeMode.Disable;
                // ユーザーIDにセットフォーカス
                ActiveControl = TxtUserId;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "【LoginForm_Load】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 「ログイン」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLogin_Click(object sender, EventArgs e)
        {
            try
            {
                CommonModule.OutPutLogFile("「ログイン」ボタンクリック");

                bool bRet = CheckUserAndPassword();
                if (!bRet)
                {
                    CommonModule.OutPutLogFile($"ログインエラー（ユーザー：{TxtUserId.Text}／パスワード：{TxtPassword.Text}）");
                    return;
                }
                CommonModule.OutPutLogFile($"ログイン成功（ユーザー：{TxtUserId.Text}）");

                TxtUserId.Text = "";
                TxtPassword.Text = "";
                TxtPassword.PasswordChar = '*';
                BtnPassword.Image = Properties.Resources.password_open;

                MenuForm form = new MenuForm();
                form.Show(this);
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnLogin_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                CommonModule.OutPutLogFile("ログイン画面からQRソータ検査アプリの終了");
                this.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【BtnEnd_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// IDとパスワードのチェック処理
        /// </summary>
        /// <returns></returns>
        private bool CheckUserAndPassword()
        {
            try
            {
                string[] sArray;
                bool bRet = false;
                foreach (var item in PubConstClass.lstUserAccount)
                {
                    sArray = item.Split(',');
                    if (sArray[0] == TxtUserId.Text)
                    {
                        if (sArray[3] == TxtPassword.Text)
                        {
                            PubConstClass.sUserId = sArray[0];
                            PubConstClass.sUserName = sArray[1];
                            PubConstClass.sUserAuthority = sArray[2];
                            PubConstClass.sUserPassword = sArray[3];
                            bRet = true;                            
                        }
                        else
                        {
                            bRet = false;
                            MessageBox.Show("パスワードが違います","警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);                                                    
                        }
                        return bRet;
                    }
                }
                if (!bRet)
                {
                    MessageBox.Show("IDが違います", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return bRet;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【CheckUserAndPassword】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 「目のアイコン」ボタン処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnPassword_Click(object sender, EventArgs e)
        {
            try
            {
                if (TxtPassword.PasswordChar.ToString() == "*")
                {
                    TxtPassword.PasswordChar = '\0';
                    BtnPassword.Image = Properties.Resources.password_close;
                }
                else
                {
                    TxtPassword.PasswordChar = '*';
                    BtnPassword.Image = Properties.Resources.password_open;
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, "【BtnPassword_Click】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 設定ファイルの存在チェック
        /// </summary>
        private bool ConfigurationFileExistenceCheck()
        {
            List<string> sCheckFileName = new List<string>() {
                                                                "QrSorterInspectionApp.def",
                                                                "NonDeliveryReasonSorting.txt",
                                                                "JobEntryList.txt",
                                                                "UserAccount.enc",
                                                              };
            try
            {
                string sNoFindFileName = "";
                string sStartPath = CommonModule.IncludeTrailingPathDelimiter(Environment.CurrentDirectory);

                foreach (string s in sCheckFileName)
                {
                    if (File.Exists(sStartPath + s) == false)
                    {
                        sNoFindFileName += s + Environment.NewLine;
                    }
                }

                if (sNoFindFileName != "")
                {
                    MessageBox.Show(sNoFindFileName, "下記の設定ファイルが見つかりません", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "【ConfigurationFileExistenceCheck】", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
