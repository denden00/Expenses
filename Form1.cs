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
using Microsoft.VisualBasic.FileIO;

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
        public readonly Dictionary<int,string> CategoryPairs = new Dictionary<int, string>
        {
            {0, "食費" },
            { 1, "日用品" },
            { 2, "携帯代" },
            { 3, "電気代" },
            { 4, "交通費" },
            { 99, "その他" },
        };

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

        /// <summary>
        /// 明細取り込みボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InputMeisai_Click(object sender, EventArgs e)
        {
            string log="";
            ErrorMessage.Visible = false;

            try
            {

                //入力チェック
                if (!(Check_Input()))
                {
                    ErrorMessage.Visible = true;
                    return;
                }

                List<Row_Meisai> meisaiList = new List<Row_Meisai>(); // 列の値を格納するリスト
                                                                      //csv抽出
                using (var inputFileStream = new FileStream(textBox2.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (TextFieldParser parser = new TextFieldParser(inputFileStream, Encoding.GetEncoding("Shift_JIS")))

                {
                    //ここでSBIかJCBかの判定を行う　↓はSBIのパターン



                    // 1行目をチェックしてSBIかJCBか判定
                    string[] row1 = parser.ReadLine().Split(',');
                    if (row1[0] == "\"1\"")
                    {
                        //SBIパターン
                        parser.TextFieldType = FieldType.Delimited;
                        parser.SetDelimiters(",");
                        while (!parser.EndOfData)
                        {
                            // CSVファイルの行をUTF-16に変換せずにそのままSplit関数で分割
                            string[] columns = parser.ReadFields();
                            string[] values = new string[columns.Length];
                            for (int i = 0; i < columns.Length; i++)
                            {
                                string utf16Value = columns[i];
                                values[i] = utf16Value;
                            }

                            // 特定の列の項目をリストに追加
                            Row_Meisai row = new Row_Meisai();
                            row.Date = values[1].Trim('"');
                            row.Detail = values[2].Trim('"');
                            row.Price = values[4].Trim('"');
                            row.Category = Check_Category(row.Detail);

                            meisaiList.Add(row);
                        }
                    }
                    else if (row1[0] == "\"\"")
                    {
                        using (StreamReader sr = new StreamReader(textBox2.Text))
                        {
                            //JCBパターン
                            //1~6行目をスキップ
                            for (int i = 0; i < 6; i++)
                            {
                                sr.ReadLine();
                            }
                            while (!sr.EndOfStream)
                            {
                                string line = sr.ReadLine();
                                string[] values = line.Split(','); // カンマ(,)で項目を区切る

                                // 特定の列の項目をリストに追加
                                Row_Meisai row = new Row_Meisai();
                                row.Date = values[2].Trim('"');
                                row.Detail = values[3].Trim('"');
                                row.Price = values[4].Trim('"');
                                row.Category = Check_Category(row.Detail);

                                meisaiList.Add(row);
                            }
                        }

                    }

                }


                // 日付の早い順にソート
                meisaiList = meisaiList.OrderBy(x => x.Date).ToList();
                // 指定した月以外のデータを削除
                string targetMonth = DateTime.Now.Year.ToString() + "/" + MonthsComboBox.Text;
                meisaiList.RemoveAll(m => !m.Date.Contains(targetMonth));

                foreach (Row_Meisai meisai in meisaiList)
                {
                    log += Check_Pair(meisai.Category) + " ： " + meisai.Detail + Environment.NewLine;
                }
                logTextBox.Text = log;

                //エクセル開いておく

                //meisaiList
            }
            catch (Exception ex)
            {
                ErrorMessage.Text = "何らかのエラーが発生しました。";
                ErrorMessage.Visible = true;
                log= ex.Message + Environment.NewLine + ex.StackTrace;
                logTextBox.Text = log;
            }
        }

        private int Check_Category(string detail)
        {
            // 正規表現パターンの配列を定義する
            string[][] regexPatterns = new string[5][];
            //食費
            regexPatterns[0] = new string[]
            {
            @"ﾊﾞﾛ-(ｽ-ﾊﾟ-)",       // バロー
            @"バロー",  // Vドラッグ
            @"ﾋﾟｱｺﾞ",                //
            @"ﾏﾙｽｱﾝｼﾞﾖｳ", // マルス
            @"ﾋﾞﾂｸﾞﾘﾌﾞ", // ビッグリブ
            @"ｺ-ﾌﾟｱｲﾁ", // コープ
            @"ﾏﾙｽｱﾝｼﾞﾖｳ", // マルス
            @"Ｖドラッグ", // Ｖドラッグ 

            };
            //日用品
            regexPatterns[1] = new string[]
            {
            @"マツモトキヨシ",       // マツモトキヨシ
            @"ﾏﾂﾓﾄｷﾖｼ",       // マツモトキヨシ
            @"ダイソー",  // ダイソー
            @"ﾀﾞｲｿｰ",  // ダイソー
            @"ドンキホーテ",                // ドン・キホーテ
            @"ﾄﾞﾝｷﾎｰﾃ",                // ドン・キホーテ
            @"ＡＭＡＺＯＮ", // Amazon
            @"Ａｍａｚｏｎ", // Amazon
            @"ドン・キホーテ", // ドン・キホーテ

            };
            //携帯
            regexPatterns[2] = new string[]
            {
            @"ｿﾌﾄﾊﾞﾝｸM",       // 携帯代
            };
            //電気
            regexPatterns[3] = new string[]
            {
            @"中部電力",       // 電気代
            };
            //交通費
            regexPatterns[4] = new string[]
            {
            @"Ｓｕｉｃａ",       // 電車チャージ分
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

        /// <summary>
        /// 入力欄に情報が入力されているかチェック、NGであればエラーメッセージを表示し処理終了
        /// </summary>
        /// <returns></returns>
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

        private string Check_Pair(int category)
        {
            if (CategoryPairs.TryGetValue(category, out string value))
            {
                return value;
            }
            else
            {
                // 存在しない数値が指定された場合の処理
                return "不明な支出";
            }
        }
    }
}
