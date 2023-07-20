using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Expenses
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void SelectButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog.FileName;
            }
        }

        private void SelectButton2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog.FileName;
            }
        }

        private void InputMeisai_Click(object sender, EventArgs e)
        {
            ErrorMessage.Visible = false;

            if (!(Check_Input()))
            {
                ErrorMessage.Visible=true;
                return;
            }

            List<Row_Meisai> meisaiList = new List<Row_Meisai>(); // 列の値を格納するリスト
            //csv抽出
            using (StreamReader sr = new StreamReader(textBox2.Text))
            {
                // 1行目をスキップ
                sr.ReadLine();

                while (!sr.EndOfStream)
                {
                    string line = sr.ReadLine();
                    string[] values = line.Split(','); // カンマ(,)で項目を区切る

                    // 特定の列の項目をリストに追加
                    Row_Meisai row = new Row_Meisai();
                    row.Date = values[1].Trim('"');
                    row.Detail = values[2].Trim('"');
                    row.Price = values[4].Trim('"');

                    meisaiList.Add(row);
                }
            }
            // 日付の早い順にソート
            meisaiList=meisaiList.OrderBy(x => x.Date).ToList();
            // 指定した月以外のデータを削除
            string targetMonth = DateTime.Now.Year.ToString() + "/" + MonthsComboBox.Text;
            meisaiList.RemoveAll(m => !m.Date.Contains(targetMonth));
        }

        private bool Check_Input()
        {
            //月チェック
            if (MonthsComboBox.SelectedItem == null)
            {
                ErrorMessage.Text = "月を選択してください";
                return false;
            }

            if (string.IsNullOrEmpty(textBox1.Text))
            {
                ErrorMessage.Text = "貼り付け先ファイルを選択してください";
                return false;
            }

            if (string.IsNullOrEmpty(textBox2.Text))
            {
                ErrorMessage.Text = "明細ファイルを選択してください";
                return false;
            }

            if (!(File.Exists(textBox1.Text)))
            {
                ErrorMessage.Text = "貼り付け先ファイルが存在しません";
                return false;
            }

            if (!(File.Exists(textBox2.Text)))
            {
                ErrorMessage.Text = "明細ファイルが存在しません";
                return false;
            }

            return true;
        }
    }
}
