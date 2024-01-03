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
using Excel=Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using Microsoft.VisualBasic.Logging;
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
        public readonly int HOKENDAI = 5;
        public readonly int OTHER = 99;
        public readonly Dictionary<int,string> CategoryPairs = new Dictionary<int, string>
        {
            {0, "食費" },
            { 1, "日用品費" },
            { 2, "携帯代" },
            { 3, "電気代" },
            { 4, "交通費" },
            { 5, "保険代" },
            { 99, "その他" },
        };

        public Form1()
        {
            InitializeComponent();
        }

        private void SelectButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.InitialDirectory = "C:\\Programming\\PersonalProject\\Expenses\\src\\Documents\\家計簿"; //デバッグ用
            openFileDialog.InitialDirectory = "C:\\Users\\Public\\Documents\\家計簿"; //本番用

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ExpensesPathTextBox.Text = openFileDialog.FileName;
            }
        }

        private void SelectButton2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.InitialDirectory = "C:\\Programming\\PersonalProject\\Expenses\\src\\Documents\\家計簿\\明細"; //デバッグ用
            openFileDialog.InitialDirectory = "C:\\Users\\Public\\Documents\\家計簿\\明細"; //本番用

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                MeisaiPathTextBox.Text = openFileDialog.FileName;
            }
        }

        /// <summary>
        /// 明細取り込みボタン押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InputMeisai_Click(object sender, EventArgs e)
        {
            EnableOff_Button();

            string log="";
            ErrorMessage.Visible = false;

            try
            {

                //入力チェック
                if (!(Check_Input()))
                {
                    ErrorMessage.Visible = true;
                    EnableOn_Button();
                    return;
                }

                List<Row_Meisai> meisaiList = new List<Row_Meisai>(); // 列の値を格納するリスト
                                                                      //csv抽出
                using (var inputFileStream = new FileStream(MeisaiPathTextBox.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
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
                        using (StreamReader sr = new StreamReader(MeisaiPathTextBox.Text))
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
                                //string[] values = line.Split(','); // カンマ(,)で項目を区切る
                                string[] values = line.Split(new string[] {"\",\""}, StringSplitOptions.None);

                                // 特定の列の項目をリストに追加
                                Row_Meisai row = new Row_Meisai();
                                row.Date = values[2].Trim();
                                row.Detail = values[3].Trim();
                                row.Price = values[4].Trim().Replace(",",""); //5,000などに対処
                                row.Category = Check_Category(row.Detail);

                                meisaiList.Add(row);
                            }
                        }

                    }

                }


                // 日付の早い順にソート
                meisaiList = meisaiList.OrderBy(x => x.Date).ToList();
                // 指定した月以外のデータを削除
                string targetYear;
                if (MonthsComboBox.Text == "12")
                {
                    targetYear = (DateTime.Now.Year - 1).ToString();
                }
                else
                {
                    targetYear = DateTime.Now.Year.ToString();
                }

                string targetMonth = targetYear + "/" + MonthsComboBox.Text;
                meisaiList.RemoveAll(m => !m.Date.Contains(targetMonth));

                if(meisaiList.Count == 0)
                {
                    ErrorMessage.Text = "該当月の明細がありません。";
                    ErrorMessage.Visible = true;
                    EnableOn_Button();
                    return;
                }

                //エクセル処理
                Post_Excel(meisaiList);
                log += "家計簿への書き込みが完了しました。" + Environment.NewLine;

                foreach (Row_Meisai meisai in meisaiList)
                {
                    log += Check_Pair(meisai.Category) + " ： " + meisai.Detail + Environment.NewLine;
                }
                logTextBox.Text = log;

                EnableOn_Button();
            }
            catch (Exception ex)
            {
                ErrorMessage.Text = "何らかのエラーが発生しました。";
                ErrorMessage.Visible = true;
                log= ex.Message + Environment.NewLine + ex.StackTrace;
                logTextBox.Text = log;
            }
        }

        private void Post_Excel(List<Row_Meisai> meisaiList)
        {
            // エクセルアプリケーションの起動
            Excel.Application excelApp = new Excel.Application();
            // ワークブックを開く
            Excel.Workbook workbook = excelApp.Workbooks.Open(ExpensesPathTextBox.Text);
            // ワークシートを開く
            Excel.Worksheet worksheet=new Excel.Worksheet();

            try
            {
                //明細行分だけ繰り返し
                foreach (Row_Meisai meisai in meisaiList)
                {
                    // ワークシートを名前で指定
                    string sheetName = Check_Pair(meisai.Category); // ワークシートの名前
                    worksheet = (Excel.Worksheet)workbook.Sheets[sheetName];

                    // 特定の文字列を検索するメソッドを呼び出し、該当するセルを見つける
                    Excel.Range searchRange = worksheet.UsedRange; // 検索範囲を全体の範囲に設定（例：A1から最終セルまで）
                    string searchString = "日付(" + int.Parse(MonthsComboBox.Text) + "月)"; // 検索する文字列
                    Excel.Range resultCell = FindCellWithText(searchRange, searchString);

                    if (resultCell != null)
                    {
                        // 該当するセルの1つ下のセルが空白でない場合、空白のセルを見つけるまで繰り返す
                        Excel.Range dateCell = worksheet.Cells[resultCell.Row + 1, resultCell.Column];
                        while (!string.IsNullOrEmpty(dateCell.Value?.ToString()))
                        {
                            resultCell = dateCell;
                            dateCell = worksheet.Cells[resultCell.Row + 1, resultCell.Column];
                        }

                        // 日付セルにデータを入力する
                        dateCell.Value = meisai.Date.Trim().Remove(0, 5);

                        // dateCellの1つ右のセル(明細)を入力する
                        Excel.Range meisaiCell = worksheet.Cells[dateCell.Row, dateCell.Column + 1];
                        meisaiCell.Value = meisai.Detail;

                        // dateCellの2つ右のセル(金額)を入力する
                        Excel.Range priceCell = worksheet.Cells[dateCell.Row, dateCell.Column + 2];
                        priceCell.Value = meisai.Price;
                    }
                    else
                    {
                        Console.WriteLine("指定した文字列が見つかりませんでした。");
                    }
                }

                // ワークブックを保存
                workbook.Save();

                // ワークブックを閉じる
                workbook.Close();

                // エクセルアプリケーションを終了
                excelApp.Quit();
            }
            finally
            {
                // 使用したオブジェクトを解放
                ReleaseObject(worksheet);
                ReleaseObject(workbook);
                ReleaseObject(excelApp);
            }

        }

        static Excel.Range FindCellWithText(Excel.Range searchRange, string searchText)
        {
            foreach (Excel.Range cell in searchRange)
            {
                if (cell.Value != null && cell.Value.ToString() == searchText)
                {
                    return cell;
                }
            }

            return null;
        }

        static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("オブジェクトの解放に失敗しました: " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private int Check_Category(string detail)
        {
            // 正規表現パターンの配列を定義する
            string[][] regexPatterns = new string[6][];
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
            @"ﾊﾞﾛ-", // Ｖドラッグ 
            @"ﾌｴﾙﾅ",                //
            @"フェルナ",                //


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
            @"AMAZON", // 
            @"ｱﾏｿﾞﾝﾌﾟﾗｲﾑｶｲﾋ", // 
            @"ドラッグスギヤマ", // 
            @"ﾄﾞﾗｯｸﾞｽｷﾞﾔﾏ", // 
            @"ＤＣＭ", // 
            @"DCM", //
            @"ﾃﾞｲ-ｼ-ｴﾑ", // 

            };
            //携帯
            regexPatterns[2] = new string[]
            {
            @"ｿﾌﾄﾊﾞﾝｸM",       // 携帯代
            @"ﾋﾞﾂｸﾞﾛ-ﾌﾞ", // 

            };
            //電気
            regexPatterns[3] = new string[]
            {
            @"中部電力",       // 電気代
            @"ﾁﾕｳﾌﾞﾃﾞﾝﾘﾖｸ", // 
            };
            //交通費
            regexPatterns[4] = new string[]
            {
            @"Ｓｕｉｃａ",       // 電車チャージ分
            @"ｽｲｶ(ｹ-ﾀｲｹﾂｻｲ)", // 
            };
            //保険代
            regexPatterns[5] = new string[]
            {
            @"ﾊﾅｻｸｾｲﾒｲ",
            @"ハナサクセイメイ",
            @"ﾈｵﾌｱ-ｽﾄｾｲﾒｲ", // 
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

            if (string.IsNullOrEmpty(ExpensesPathTextBox.Text))
            {
                ErrorMessage.Text = "貼り付け先ファイルを選択してください";
                return false;
            }

            if (string.IsNullOrEmpty(MeisaiPathTextBox.Text))
            {
                ErrorMessage.Text = "明細ファイルを選択してください";
                return false;
            }

            if (!(File.Exists(ExpensesPathTextBox.Text)))
            {
                ErrorMessage.Text = "貼り付け先ファイルが存在しません";
                return false;
            }

            if (!(File.Exists(MeisaiPathTextBox.Text)))
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

        private void EnableOn_Button()
        {
            //処理完了後はボタンを戻す
            InputMeisai.BackColor = Color.SkyBlue;
            InputMeisai.Enabled = true;
        }

        private void EnableOff_Button() 
        {
            //処理中はボタンをグレーアウト＆無効化
            InputMeisai.BackColor = Color.Gray;
            InputMeisai.Enabled = false;
        }
    }
}
