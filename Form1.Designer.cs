namespace Expenses
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.InputMeisai = new System.Windows.Forms.Button();
            this.ExpensesPathTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.MeisaiPathTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.MonthsComboBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SelectButton1 = new System.Windows.Forms.Button();
            this.SelectButton2 = new System.Windows.Forms.Button();
            this.ErrorMessage = new System.Windows.Forms.Label();
            this.logTextBox = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // InputMeisai
            // 
            this.InputMeisai.BackColor = System.Drawing.Color.SkyBlue;
            this.InputMeisai.Location = new System.Drawing.Point(34, 164);
            this.InputMeisai.Name = "InputMeisai";
            this.InputMeisai.Size = new System.Drawing.Size(92, 31);
            this.InputMeisai.TabIndex = 0;
            this.InputMeisai.Text = "取り込み";
            this.InputMeisai.UseVisualStyleBackColor = false;
            this.InputMeisai.Click += new System.EventHandler(this.InputMeisai_Click);
            // 
            // ExpensesPathTextBox
            // 
            this.ExpensesPathTextBox.Location = new System.Drawing.Point(34, 75);
            this.ExpensesPathTextBox.Name = "ExpensesPathTextBox";
            this.ExpensesPathTextBox.Size = new System.Drawing.Size(398, 19);
            this.ExpensesPathTextBox.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(34, 57);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "家計簿ファイル";
            // 
            // MeisaiPathTextBox
            // 
            this.MeisaiPathTextBox.Location = new System.Drawing.Point(34, 127);
            this.MeisaiPathTextBox.Name = "MeisaiPathTextBox";
            this.MeisaiPathTextBox.Size = new System.Drawing.Size(398, 19);
            this.MeisaiPathTextBox.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(34, 109);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(63, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "明細ファイル";
            // 
            // MonthsComboBox
            // 
            this.MonthsComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.MonthsComboBox.FormattingEnabled = true;
            this.MonthsComboBox.Items.AddRange(new object[] {
            "01",
            "02",
            "03",
            "04",
            "05",
            "06",
            "07",
            "08",
            "09",
            "10",
            "11",
            "12"});
            this.MonthsComboBox.Location = new System.Drawing.Point(89, 20);
            this.MonthsComboBox.MaxDropDownItems = 13;
            this.MonthsComboBox.Name = "MonthsComboBox";
            this.MonthsComboBox.Size = new System.Drawing.Size(63, 20);
            this.MonthsComboBox.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(32, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "月選択";
            // 
            // SelectButton1
            // 
            this.SelectButton1.Location = new System.Drawing.Point(445, 75);
            this.SelectButton1.Name = "SelectButton1";
            this.SelectButton1.Size = new System.Drawing.Size(75, 23);
            this.SelectButton1.TabIndex = 4;
            this.SelectButton1.Text = "ファイル選択";
            this.SelectButton1.UseVisualStyleBackColor = true;
            this.SelectButton1.Click += new System.EventHandler(this.SelectButton1_Click);
            // 
            // SelectButton2
            // 
            this.SelectButton2.Location = new System.Drawing.Point(445, 127);
            this.SelectButton2.Name = "SelectButton2";
            this.SelectButton2.Size = new System.Drawing.Size(75, 23);
            this.SelectButton2.TabIndex = 4;
            this.SelectButton2.Text = "ファイル選択";
            this.SelectButton2.UseVisualStyleBackColor = true;
            this.SelectButton2.Click += new System.EventHandler(this.SelectButton2_Click);
            // 
            // ErrorMessage
            // 
            this.ErrorMessage.AutoSize = true;
            this.ErrorMessage.Font = new System.Drawing.Font("MS UI Gothic", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.ErrorMessage.ForeColor = System.Drawing.Color.Red;
            this.ErrorMessage.Location = new System.Drawing.Point(148, 164);
            this.ErrorMessage.Name = "ErrorMessage";
            this.ErrorMessage.Size = new System.Drawing.Size(142, 22);
            this.ErrorMessage.TabIndex = 5;
            this.ErrorMessage.Text = "エラーメッセージ";
            this.ErrorMessage.Visible = false;
            // 
            // logTextBox
            // 
            this.logTextBox.Location = new System.Drawing.Point(36, 211);
            this.logTextBox.Name = "logTextBox";
            this.logTextBox.ReadOnly = true;
            this.logTextBox.Size = new System.Drawing.Size(417, 335);
            this.logTextBox.TabIndex = 6;
            this.logTextBox.Text = "使い方\n①”月選択”から月を選んでください\n②家計簿ファイルを選択してください\n③取り込みたい明細ファイルを選択してください\n④”取り込み”ボタンを押してください" +
    "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(562, 583);
            this.Controls.Add(this.logTextBox);
            this.Controls.Add(this.ErrorMessage);
            this.Controls.Add(this.SelectButton2);
            this.Controls.Add(this.SelectButton1);
            this.Controls.Add(this.MonthsComboBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.MeisaiPathTextBox);
            this.Controls.Add(this.ExpensesPathTextBox);
            this.Controls.Add(this.InputMeisai);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button InputMeisai;
        private System.Windows.Forms.TextBox ExpensesPathTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox MeisaiPathTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox MonthsComboBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button SelectButton1;
        private System.Windows.Forms.Button SelectButton2;
        private System.Windows.Forms.Label ErrorMessage;
        private System.Windows.Forms.RichTextBox logTextBox;
    }
}

