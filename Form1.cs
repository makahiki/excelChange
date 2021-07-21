using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Npgsql;


namespace エクセル変換3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataSet ds;

        Class1 select = new Class1();
        Class1 tran = new Class1();

        enum Col
        {
            項目名,
            形式,
            出力対象,
            合計対象
        }

        #region イベント

        /// <summary>
        /// ファイル選択ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = SelctFile();
            ReadFile(fileName); 
        }

        /// <summary>
        /// テンプレ登録ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtFile.Text))
            {
                MessageBox.Show("ファイルを選択して下さい。");
                return;
            }
            if (String.IsNullOrEmpty(comboBox1.Text))
            {
                MessageBox.Show("テンプレートを選択または入力して下さい。");
                return;
            }
            AddUpTemplate();
        }

        /// <summary>
        /// テンプレ選択ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtFile.Text))
            {
                MessageBox.Show("ファイルを選択して下さい。");
                return;
            }

            SelectTemplate();
        }

        /// <summary>
        /// エクセル出力ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(txtFile.Text))
            {
                MessageBox.Show("ファイルを選択して下さい。");
                return;
            }

            DataTable dt = ds.Tables["TestTable"];
            ExportExcel(dt);
        }

        /// <summary>
        /// 出力表示チェックボックス
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            CheckOutPut();
        }

        #endregion

        #region メソッド

        /// <summary>
        /// ファイル選択
        /// </summary>
        /// <returns></returns>
        private string SelctFile()
        {
            DialogResult ret;
            string fileName;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "ダイアログボックスのサンプル";
                openFileDialog.CheckFileExists = true;

                ret = openFileDialog.ShowDialog();

                fileName = openFileDialog.FileName;
                txtFile.Text = openFileDialog.SafeFileName;
                //txtFile.Text = fileName;

                return fileName;
            }
        }

        /// <summary>
        /// ファイル読み込み
        /// </summary>
        /// <param name="fileName"></param>
        private void ReadFile(string fileName)
        {
            try
            {
                ds = new DataSet();
                DataTable dt = new DataTable("TestTable");
                string path = fileName;
                StreamReader sr = new StreamReader(path, Encoding.GetEncoding("Shift_JIS"));

                dataGridView1.Columns.Clear();
                dataGridView2.Columns.Clear();

                dt.Columns.Clear();
                dt.Rows.Clear();

                string linebuf = sr.ReadLine();

                string[] colArr;

                if (txtFile.Text.EndsWith("tsv"))
                {
                    colArr = linebuf.Split('\t');
                }
                else if (txtFile.Text.EndsWith("csv"))
                {
                    colArr = linebuf.Split(',');
                }
                else
                {
                    MessageBox.Show("tsvかcsvを選択して下さい。");
                    txtFile.Text = "";
                    return;
                }

                dataGridView2.Columns.Add("item_name", "項目名");

                foreach (string buf in colArr)
                {
                    dt.Columns.Add(buf);
                    dataGridView2.Rows.Add(buf);
                }

                while (linebuf != null)
                {
                    linebuf = sr.ReadLine();

                    if (linebuf == null)
                    {
                        break;
                    }

                    string[] rowArr;

                    if (txtFile.Text.EndsWith("tsv"))
                    {
                        rowArr = linebuf.Split('\t');
                        dt.Rows.Add(rowArr);
                    }
                    else if (txtFile.Text.EndsWith("csv"))
                    {
                        rowArr = linebuf.Split(',');
                        dt.Rows.Add(rowArr);
                    }
                }

                sr.Close();

                AddCmb();

                GetCmb();

                dataGridView1.DataSource = dt;

                ds.Tables.Add(dt);
            }
            catch
            {
                return;
            }
        }

        /// <summary>
        /// コンボボックス表示
        /// </summary>
        private void AddCmb()
        {
            DataGridViewComboBoxColumn cmbColumn = new DataGridViewComboBoxColumn();
            cmbColumn.Items.Add("文字列");
            cmbColumn.Items.Add("数値");
            cmbColumn.Items.Add("数値（小数点2桁）");
            cmbColumn.Items.Add("日付");

            cmbColumn.HeaderText = "形式";
            dataGridView2.Columns.Add(cmbColumn);

            DataGridViewCheckBoxColumn outColumn = new DataGridViewCheckBoxColumn();
            outColumn.HeaderText = "出力対象";
            dataGridView2.Columns.Add(outColumn);

            DataGridViewCheckBoxColumn totalColumn = new DataGridViewCheckBoxColumn();
            totalColumn.HeaderText = "合計対象";
            dataGridView2.Columns.Add(totalColumn);

            //最終行非表示
            dataGridView1.AllowUserToAddRows = false;
            dataGridView2.AllowUserToAddRows = false;

            //デフォルト
            int dgvRowCnt = dataGridView2.Rows.Count;
            for (int i = 0; i < dgvRowCnt; i++)
            {
                dataGridView2.Rows[i].Cells[(int)Col.形式].Value = "文字列";
                dataGridView2.Rows[i].Cells[(int)Col.出力対象].Value = true;
                dataGridView2.Rows[i].Cells[(int)Col.合計対象].Value = false;
            }
        }

        /// <summary>
        /// テンプレートコンボ表示
        /// </summary>
        private void GetCmb()
        {
            string sql;
            sql = "";
            sql += " select distinct";
            sql += " template_id,";
            sql += " template_name";
            sql += " from save_template";
            sql += " order by template_id asc";

            //sql実行
            //DataTable dt = SelectSpl(sql);

            //Class1 select = new Class1();
            DataTable dt = select.SelectSpl(sql);

            DataTable dtCombo = new DataTable();
            dtCombo.Columns.Add("template_id");
            dtCombo.Columns.Add("template_name");

            DataRow dtRowCombo;

            foreach (DataRow row in dt.Rows)
            {
                dtRowCombo = dtCombo.NewRow();

                dtRowCombo["template_id"] = row["template_id"];
                dtRowCombo["template_name"] = row["template_name"];
                dtCombo.Rows.Add(dtRowCombo);
            }

            comboBox1.DataSource = dtCombo;
            comboBox1.DisplayMember = ("template_name");
            comboBox1.ValueMember = ("template_id");
        }

        /// <summary>
        /// 出力表示
        /// </summary>
        private void CheckOutPut()
        {
            Boolean cheakStatus;
            int dgvRowCnt = dataGridView2.Rows.Count;

            if(checkBox1.Checked == true)
            {
                for(int i = 0; i < dgvRowCnt; i++)
                {
                    cheakStatus = System.Convert.ToBoolean(dataGridView2.Rows[i].Cells[(int)Col.出力対象].Value);

                    if (cheakStatus == true)
                    {
                        dataGridView1.Columns[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Columns[i].Visible = false;
                    }
                }
            }
            else
            {
                for(int i = 0; i < dgvRowCnt; i++)
                {
                    dataGridView1.Columns[i].Visible = true;
                }
            }
        }

        /// <summary>
        /// 数字⇒文字列
        /// </summary>
        /// <param name="columnNo"></param>
        /// <returns></returns>
        private string Toalphabet(int columnNo)
        {
            string alphabet = "ZABCDEFGHIJKLMNOPQRSTUVWXY";
            string columnStr = string.Empty;
            int m = 0;
            do
            {
                m = columnNo % 26;
                columnStr = alphabet[m] + columnStr;
                columnNo = columnNo / 26;

            } while (0 < columnNo && m != 0);

            return columnStr;
        }

        /// <summary>
        /// エクセル出力
        /// </summary>
        /// <param name="dt"></param>
        public void ExportExcel(DataTable dt)
        {
            dynamic xlApp = null;
            dynamic xlBooks = null;
            dynamic xlBook = null;
            dynamic xlSheet = null;
            dynamic xlCells = null;
            dynamic xlRange = null;
            dynamic xlRange2 = null;
            dynamic xlCellStart = null;
            dynamic xlCellEnd = null;
            try
            {
                xlApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                xlBooks = xlApp.Workbooks;
                xlBook = xlBooks.Add;
                xlSheet = xlBook.WorkSheets(1);
                xlCells = xlSheet.Cells;

                DataColumn dc;
                object[,] columnData = new object[dt.Rows.Count, 1];
                int row = 1;
                int col = 1;
                int col2 = 1;
                int rowCount;
                string sumExcel;
                string changeExcelRow;
                Boolean cheakStatus;
                Boolean cheakStatus2;

                for (col = 1; (col <= dt.Columns.Count); col++)
                {
                    cheakStatus = System.Convert.ToBoolean(dataGridView2.Rows[col-1].Cells[(int)Col.出力対象].Value);
                    if (cheakStatus == true)
                    {
                        row = 1;
                        dc = dt.Columns[(col - 1)];
                        // ヘッダー行の出力
                        xlCells[row, col2].value2 = dc.ColumnName;
                        row++;
                        // 列データを配列に格納
                        for (int i = 0; (i <= (dt.Rows.Count - 1)); i++)
                        {
                            columnData[i, 0] = string.Format("{0}", dt.Rows[i][(col - 1)]);
                        }

                        xlCellStart = xlCells[row, col2];
                        xlCellEnd = xlCells[(row + (dt.Rows.Count - 1)), col2];
                        xlRange = xlSheet.Range(xlCellStart, xlCellEnd);

                        //フォーマット
                        switch (dataGridView2.Rows[col - 1].Cells[1].Value)
                        {
                            case "文字列":
                                xlRange.NumberFormatLocal = "@";
                                break;
                            case "日付":
                                xlRange.NumberFormatLocal = "yyyy/mm/dd";
                                break;
                            case "数値":
                                xlRange.NumberFormatLocal = "#,##0";
                                break;
                            case "数値（小数点2桁）":
                                xlRange.NumberFormatLocal = "#,##0.00";
                                break;
                        }

                        xlRange.value2 = columnData;

                        cheakStatus2 = System.Convert.ToBoolean(dataGridView2.Rows[col - 1].Cells[(int)Col.合計対象].Value);

                        //合計出力
                        if (cheakStatus2 == true)
                        {
                            changeExcelRow = Toalphabet(col2);
                            rowCount = dt.Rows.Count + 1;
                            sumExcel = "=SUM(" + changeExcelRow + "2:" + changeExcelRow + rowCount + ")";
                            xlRange2 = xlSheet.Cells(dt.Rows.Count + 2, changeExcelRow);
                            xlRange2.value = sumExcel;
                        }
                        col2 += 1;
                    }
                }
                xlCells.EntireColumn.AutoFit();
                xlRange = xlSheet.UsedRange;
                xlRange.Borders.LineStyle = 1;  // xlContinuous
                xlApp.Visible = true;
            }
            catch
            {
                xlApp.DisplayAlerts = false;
                xlApp.Quit();
                throw;
            }
            finally
            {
                if (xlCellStart != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCellStart);
                if (xlCellEnd != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCellEnd);
                if (xlRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                if (xlRange2 != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange2);
                if (xlCells != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCells);
                if (xlSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                if (xlBook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                if (xlBooks != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                if (xlApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                GC.Collect();
            }
        }

        /// <summary>
        /// テンプレートid取得
        /// </summary>
        /// <returns></returns>
        private int GetTemplateId()
        {
            string sql;
            sql = "";
            sql += " select distinct";
            sql += " template_id,";
            sql += " template_name";
            sql += " from save_template";
            sql += " order by template_id desc";

            //sql実行
            //DataTable dt = SelectSpl(sql);

            //Class1 select = new Class1();
            DataTable dt = select.SelectSpl(sql);

            int rowCount = dt.Rows.Count;
            int templateId;

            if (rowCount == 0)
            {
                templateId = 1;
            }
            else
            {
                templateId = int.Parse(dt.Rows[0]["template_id"].ToString()) + 1;
            }
            return templateId;
        }

        /// <summary>
        /// テンプレート選択
        /// </summary>
        private void SelectTemplate()
        {
            string sql;
            sql = "";
            sql += " select";
            sql += " * from";
            sql += " save_template";
            sql += " where template_id = " + comboBox1.SelectedValue;
            sql += " order by id asc";

            //sql実行
            //DataTable dt = SelectSpl(sql);

            //Class1 select = new Class1();
            DataTable dt = select.SelectSpl(sql);

            int dgvRowCnt = dataGridView2.Rows.Count;

            for (int i = 0; i < dgvRowCnt; i++)
            {
                var originName = dataGridView1.Columns[i].HeaderCell.Value;
                 

                foreach (DataRow row in dt.Rows)
                {
                    var dtName = row["origin_name"];
                    if (originName.Equals(dtName))
                    {
                        dataGridView2.Rows[i].Cells[0].Value = row["item_name"];
                        dataGridView2.Rows[i].Cells[1].Value = row["format"];
                        dataGridView2.Rows[i].Cells[2].Value = row["output_target"];
                        dataGridView2.Rows[i].Cells[3].Value = row["total_target"];
                    }
                }
            }
        }

        /// <summary>
        /// テンプレート登録or更新
        /// </summary>
        private void AddUpTemplate()
        {
            string sql;
            sql = "";
            sql += " select distinct";
            sql += " template_id,";
            sql += " template_name";
            sql += " from save_template";
            sql += " where template_name = '" + comboBox1.Text + "'";

            //sql実行
            //DataTable dt = SelectSpl(sql);

            //Class1 select = new Class1();
            DataTable dt = select.SelectSpl(sql);

            int rowCount = dt.Rows.Count;
            int templateId = GetTemplateId();
            int dgvRowCnt = dataGridView2.Rows.Count;

            if (rowCount == 0)
            {
                for(int i = 0; i < dgvRowCnt; i++)
                {
                    var originName = dataGridView1.Columns[i].HeaderCell.Value;
                    var itemName = dataGridView2.Rows[i].Cells[(int)Col.項目名].Value;
                    var format = dataGridView2.Rows[i].Cells[(int)Col.形式].Value;
                    var outPut = dataGridView2.Rows[i].Cells[(int)Col.出力対象].Value;
                    var total = dataGridView2.Rows[i].Cells[(int)Col.合計対象].Value;

                    sql = "";
                    sql += " insert";
                    sql += " into save_template";
                    sql += " (template_id, template_name, origin_name, item_name, format, output_target, total_target)";
                    sql += " values";
                    sql += " (" + templateId + ", '" + comboBox1.Text + "', '" + originName + "', '" + itemName + "', '" + format + "'," + outPut + "," + total + ")";

                    // sql実行
                    //TranSpl(sql);

                    //Class1 tran = new Class1();
                    tran.TranSpl(sql);
                }
                MessageBox.Show("登録完了");
                GetCmb();
                comboBox1.SelectedValue = templateId; 
            }
            else
            {
                for (int i = 0; i < dgvRowCnt; i++)
                {
                    var originName = dataGridView1.Columns[i].HeaderCell.Value;
                    var itemName = dataGridView2.Rows[i].Cells[(int)Col.項目名].Value;
                    var format = dataGridView2.Rows[i].Cells[(int)Col.形式].Value;
                    var outPut = dataGridView2.Rows[i].Cells[(int)Col.出力対象].Value;
                    var total = dataGridView2.Rows[i].Cells[(int)Col.合計対象].Value;

                    sql = "";
                    sql += " update save_template";
                    sql += " set item_name ='" + itemName + "',";
                    sql += " format = '" + format + "',";
                    sql += " output_target = " + outPut + ",";
                    sql += " total_target = " + total;
                    sql += " where template_name = '" + comboBox1.Text + "'";
                    sql += " and origin_name = '" + originName + "'";

                    //sql実行
                    //TranSpl(sql);

                    //Class1 tran = new Class1();
                    tran.TranSpl(sql);
                }
                MessageBox.Show("更新完了");
            }
        }

        #endregion

        #region sql実行

        ///// <summary>
        ///// update,insert用sql実行
        ///// </summary>
        ///// <param name="sql"></param>
        //private void TranSpl(string sql)
        //{
        //    // 接続文字列
        //    var connString = "Server=localhost;Port=5432;Username=postgres;Password=postgres;Database=vending_machine2";

        //    using (var conn = new NpgsqlConnection(connString))
        //    {
        //        conn.Open();
        //        using (var transaction = conn.BeginTransaction())
        //        {
        //            var command = new NpgsqlCommand(sql, conn);
        //            command.Parameters.Add(new NpgsqlParameter("p", DbType.Int32) { Value = 123 });

        //            try
        //            {
        //                command.ExecuteNonQuery();
        //                transaction.Commit();
        //            }
        //            catch (NpgsqlException)
        //            {
        //                transaction.Rollback();
        //                throw;
        //            }
        //        }
        //    }
        //}

        ///// <summary>
        ///// select用sql実行
        ///// </summary>
        ///// <param name="sql"></param>
        ///// <returns></returns>
        //private DataTable SelectSpl(string sql)
        //{
        //    // 接続文字列
        //    var connString = "Server=localhost;Port=5432;Username=postgres;Password=postgres;Database=vending_machine2";

        //    using (var conn = new NpgsqlConnection(connString))
        //    {
        //        conn.Open();

        //        var dataAdapter = new NpgsqlDataAdapter(sql, conn);

        //        DataTable dt = new DataTable();
        //        dataAdapter.Fill(dt);

        //        return dt;
        //    }
        //}

        #endregion 

    }
}
