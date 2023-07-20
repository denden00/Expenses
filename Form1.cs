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
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace Expenses
{
    public partial class Form1 : Form
    {
        public readonly int SYOKUHI= 0;
        public readonly int HITIYOUHINDAI = 1;
        public readonly int KEITAIDAI = 2;
        public readonly int DENKIDAI = 3;
        public readonly int KOUTUUHI = 4;
        public readonly int OTHER = 99;

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
                //ここでSBIかJCBかの判定を行う　↓はSBIのパターン

                //SBIパターン
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
                    row.Category = Check_Category(row.Detail);

                    meisaiList.Add(row);
                }

                //JCBパターン（※SBIとはIFでつなげる）
            }


            // 日付の早い順にソート
            meisaiList=meisaiList.OrderBy(x => x.Date).ToList();
            // 指定した月以外のデータを削除
            string targetMonth = DateTime.Now.Year.ToString() + "/" + MonthsComboBox.Text;
            meisaiList.RemoveAll(m => !m.Date.Contains(targetMonth));

            //エクセル開いておく

            //meisaiList

        }

        private int Check_Category(string detail)
        {
            // 正規表現パターンの配列を定義する
            string[][] regexPatterns = new string[5][];
            //食費
            regexPatterns[0] = new string[]
            {
            @"ﾊﾞﾛ-(ｽ-ﾊﾟ-)",       // バロー
            @"^[a-z]+(_[a-z]+)*$",  // ドミー
            @"^\d+$",                // アミカ
            @"^[a-zA-Z0-9]+@[a-zA-Z0-9]+\.[a-zA-Z]+$", // ピアゴ
            };
            //日用品
            regexPatterns[1] = new string[]
            {
            @"ﾊﾞﾛ-(ｽ-ﾊﾟ-)",       // バロー
            @"^[a-z]+(_[a-z]+)*$",  // ドミー
            @"^\d+$",                // アミカ
            @"^[a-zA-Z0-9]+@[a-zA-Z0-9]+\.[a-zA-Z]+$", // ピアゴ
            };
            //携帯
            regexPatterns[2] = new string[]
            {
            @"ﾊﾞﾛ-(ｽ-ﾊﾟ-)",       // バロー
            @"^[a-z]+(_[a-z]+)*$",  // ドミー
            @"^\d+$",                // アミカ
            @"^[a-zA-Z0-9]+@[a-zA-Z0-9]+\.[a-zA-Z]+$", // ピアゴ
            };
            //電気
            regexPatterns[3] = new string[]
            {
            @"ﾊﾞﾛ-(ｽ-ﾊﾟ-)",       // バロー
            @"^[a-z]+(_[a-z]+)*$",  // ドミー
            @"^\d+$",                // アミカ
            @"^[a-zA-Z0-9]+@[a-zA-Z0-9]+\.[a-zA-Z]+$", // ピアゴ
            };
            //交通費
            regexPatterns[4] = new string[]
            {
            @"ﾊﾞﾛ-(ｽ-ﾊﾟ-)",       // バロー
            @"^[a-z]+(_[a-z]+)*$",  // ドミー
            @"^\d+$",                // アミカ
            @"ｿﾌﾄﾊﾞﾝｸM", // ピアゴ
            };

            int rows = regexPatterns.Length; // 行数を取得
            
            for (int i = 0; i < rows; i++)
            {
                int columns = regexPatterns[i].Length; // 列数を取得
                for (int j = 0; j < columns; j++)
                {
                    //if (Regex.IsMatch(detail, regexPatterns[i][j]))
                    if (detail.Contains(regexPatterns[i][j]))
                    {
                        return i;
                    }
                }
            }
                    return 99;
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
