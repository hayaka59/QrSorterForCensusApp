using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QrSorterInspectionApp
{
    public partial class ErrorMessageForm : Form
    {
        private static ErrorMessageForm _singleInstance = new ErrorMessageForm();

        // シングルトンパターン
        public ErrorMessageForm()
        {
            InitializeComponent();

            // フォーム表示位置指定（画面の中央に表示）
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point((1920 - this.Width) / 2, (1080 - this.Height) / 2);
            // フォームスタイル設定
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.WindowState = FormWindowState.Normal;
            // ウィンドウサイズ設定
            this.Width = 830;
            this.Height = 281;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static ErrorMessageForm GetInstance()
        {
            return _singleInstance;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sErrorData"></param>
        public ErrorMessageForm(string sErrorData)
        {
            string[] sArray = sErrorData.Split(',');

            try
            {
                // エラー内容の表示
                LblTitle.Text = sArray[1];
                LblErrorMessage.Text = sArray[2];
                LblErrorNumber.Text = "エラーNo." + sArray[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー【ErrorMessageForm】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sErrorData"></param>
        public void SetMessage(string sErrorData)
        {
            string[] sArray = sErrorData.Split(',');

            try
            {
                // エラー内容の表示
                LblTitle.Text = sArray[1];
                LblErrorMessage.Text = sArray[2];
                LblErrorNumber.Text = "エラーNo." + sArray[0];

                if (PubConstClass.bIsErrorMessage)
                {
                    LblErrorNumber.Text = "エラーNo." + sArray[0];
                    this.Text = "エラーメッセージ";
                    LblTitle.BackColor = Color.Salmon;
                }
                else
                {
                    LblErrorNumber.Text = "No." + sArray[0];
                    this.Text = "確認メッセージ";
                    LblTitle.BackColor = Color.Green;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー【SetMessage】", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnRelease_Click(object sender, EventArgs e)
        {
            // エラーリセットコマンド送信（d）            
            QrSorterInspectionForm.SendResetCommand();
            
            PubConstClass.bIsOpenErrorMessage = false;
            this.Hide();
        }

        /// <summary>
        /// 「✕」ボタンの無効化
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ErrorMessageForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //LblTitle.Text = "";
            e.Cancel = true;
        }

        private void ErrorMessageForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            //LblTitle.Text = "";
        }
    }
}
