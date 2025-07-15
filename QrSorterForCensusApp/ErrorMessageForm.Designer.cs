namespace QrSorterInspectionApp
{
    partial class ErrorMessageForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Panel1 = new System.Windows.Forms.Panel();
            this.LblErrorNumber = new System.Windows.Forms.Label();
            this.BtnRelease = new System.Windows.Forms.Button();
            this.LblErrorMessage = new System.Windows.Forms.Label();
            this.LblTitle = new System.Windows.Forms.Label();
            this.Panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Panel1
            // 
            this.Panel1.Controls.Add(this.BtnRelease);
            this.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.Panel1.Location = new System.Drawing.Point(0, 168);
            this.Panel1.Name = "Panel1";
            this.Panel1.Size = new System.Drawing.Size(814, 74);
            this.Panel1.TabIndex = 8;
            // 
            // LblErrorNumber
            // 
            this.LblErrorNumber.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.LblErrorNumber.BackColor = System.Drawing.Color.White;
            this.LblErrorNumber.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblErrorNumber.ForeColor = System.Drawing.Color.Blue;
            this.LblErrorNumber.Location = new System.Drawing.Point(3, 46);
            this.LblErrorNumber.Name = "LblErrorNumber";
            this.LblErrorNumber.Size = new System.Drawing.Size(184, 20);
            this.LblErrorNumber.TabIndex = 1;
            this.LblErrorNumber.Text = "LblErrorNumber";
            this.LblErrorNumber.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // BtnRelease
            // 
            this.BtnRelease.Font = new System.Drawing.Font("メイリオ", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.BtnRelease.Image = global::QrSorterInspectionApp.Properties.Resources.check;
            this.BtnRelease.Location = new System.Drawing.Point(266, 10);
            this.BtnRelease.Name = "BtnRelease";
            this.BtnRelease.Size = new System.Drawing.Size(300, 47);
            this.BtnRelease.TabIndex = 0;
            this.BtnRelease.Text = "OK";
            this.BtnRelease.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.BtnRelease.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnRelease.UseVisualStyleBackColor = true;
            this.BtnRelease.Click += new System.EventHandler(this.BtnRelease_Click);
            // 
            // LblErrorMessage
            // 
            this.LblErrorMessage.BackColor = System.Drawing.Color.White;
            this.LblErrorMessage.Dock = System.Windows.Forms.DockStyle.Top;
            this.LblErrorMessage.Font = new System.Drawing.Font("メイリオ", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.LblErrorMessage.ForeColor = System.Drawing.Color.Red;
            this.LblErrorMessage.Location = new System.Drawing.Point(0, 43);
            this.LblErrorMessage.Name = "LblErrorMessage";
            this.LblErrorMessage.Size = new System.Drawing.Size(814, 112);
            this.LblErrorMessage.TabIndex = 7;
            this.LblErrorMessage.Text = "LblErrorMessage";
            this.LblErrorMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // LblTitle
            // 
            this.LblTitle.BackColor = System.Drawing.Color.Salmon;
            this.LblTitle.Dock = System.Windows.Forms.DockStyle.Top;
            this.LblTitle.Font = new System.Drawing.Font("メイリオ", 18F);
            this.LblTitle.ForeColor = System.Drawing.Color.White;
            this.LblTitle.Location = new System.Drawing.Point(0, 0);
            this.LblTitle.Name = "LblTitle";
            this.LblTitle.Size = new System.Drawing.Size(814, 43);
            this.LblTitle.TabIndex = 6;
            this.LblTitle.Text = "LblTitle";
            this.LblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ErrorMessageForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(814, 242);
            this.ControlBox = false;
            this.Controls.Add(this.LblErrorNumber);
            this.Controls.Add(this.Panel1);
            this.Controls.Add(this.LblErrorMessage);
            this.Controls.Add(this.LblTitle);
            this.Font = new System.Drawing.Font("メイリオ", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(5, 6, 5, 6);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ErrorMessageForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "エラーメッセージ";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ErrorMessageForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ErrorMessageForm_FormClosed);
            this.Panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Panel Panel1;
        internal System.Windows.Forms.Label LblErrorNumber;
        internal System.Windows.Forms.Button BtnRelease;
        internal System.Windows.Forms.Label LblErrorMessage;
        internal System.Windows.Forms.Label LblTitle;
    }
}