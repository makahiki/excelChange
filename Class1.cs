using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Npgsql;

namespace エクセル変換3
{
    class Class1
    {

        /// <summary>
        /// update,insert用sql実行
        /// </summary>
        /// <param name="sql"></param>
        public  void TranSpl(string sql)
        {
            // 接続文字列
            var connString = "Server=localhost;Port=5432;Username=postgres;Password=postgres;Database=vending_machine2";

            using (var conn = new NpgsqlConnection(connString))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    var command = new NpgsqlCommand(sql, conn);
                    command.Parameters.Add(new NpgsqlParameter("p", DbType.Int32) { Value = 123 });

                    try
                    {
                        command.ExecuteNonQuery();
                        transaction.Commit();
                    }
                    catch (NpgsqlException)
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        /// <summary>
        /// select用sql実行
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public DataTable SelectSpl(string sql)
        {
            // 接続文字列
            var connString = "Server=localhost;Port=5432;Username=postgres;Password=postgres;Database=vending_machine2";

            using (var conn = new NpgsqlConnection(connString))
            {
                conn.Open();

                var dataAdapter = new NpgsqlDataAdapter(sql, conn);

                DataTable dt = new DataTable();
                dataAdapter.Fill(dt);

                return dt;
            }
        }

        ////Implements IDisposable;
        //string conStr = ("Server=localhost;Port=5432;User Id=postgres;Password=postgres;Database=vending_machine2;");
        //NpgsqlConnection sqlCon; //= new NpgsqlConnection();
        //NpgsqlTransaction sqlTrn; //= new NpgsqlTransaction();
        //NpgsqlCommand sqlCmd; //= new NpgsqlCommand();
        //NpgsqlDataAdapter sqlAdp; //= new NpgsqlDataAdapter();

        ////public void Sub New()
        ////{

        ////}

        ////DB接続開始
        //public void open()
        //{
        //    if(sqlCon == null)
        //    {
        //        sqlCon = new NpgsqlConnection(conStr);
        //        sqlCon.Open();
        //    }
        //}

        ////全てのオブジェクトを破棄し、DB接続を終了
        //public void close()
        //{
        //    if(!(sqlAdp == null))
        //    {
        //        sqlAdp.Dispose();
        //        sqlAdp = null;
        //    }
        //    if (!(sqlCmd == null))
        //    {
        //        sqlCmd.Dispose();
        //        sqlCmd = null;
        //    }
        //    if (!(sqlTrn == null))
        //    {
        //        sqlTrn.Dispose();
        //        sqlTrn = null;
        //    }
        //    if (!(sqlCon == null))
        //    {
        //        sqlCon.Dispose();
        //        sqlCon = null;
        //    }
        //}

        ////トランザクション開始
        //public void trnStart()
        //{
        //    if(sqlTrn == null)
        //    {
        //        sqlTrn = sqlCon.BeginTransaction();
        //    }
        //}

        ////トランザクションコミット
        //public void commit()
        //{
        //    if (!(sqlTrn == null))
        //    {
        //        sqlTrn.Commit();
        //    }
        //}

        ////トランザクションロールバック
        //public void rollback()
        //{
        //    if (!(sqlTrn == null))
        //    {
        //        sqlTrn.Rollback();
        //    }
        //}

        ///// <summary>
        ///// トランザクションを伴わないSQLを実行(主にSELECT文)
        ///// </summary>
        ///// <param name="sql"></param>
        ///// <returns></returns>
        //public DataTable getDtSql(string sql)
        //{
        //    //結果を格納するDataTableを宣言
        //    DataTable returnDt = new DataTable();

        //    try
        //    {
        //        sqlCmd = new NpgsqlCommand(sql, sqlCon);
        //        sqlAdp = new NpgsqlDataAdapter(sqlCmd);
        //        sqlAdp.Fill(returnDt);
        //    }   
        //    catch (Exception e)
        //    {
        //        throw;
        //    }
        //    return returnDt;
        //}

        ///// <summary>
        ///// トランザクションを伴うSQLを実行(主にINSERT,UPDATE,DELETE文)
        ///// </summary>
        ///// <param name="sql"></param>
        //public void executeSql(string sql)
        //{
        //    try
        //    {
        //        sqlCmd = new NpgsqlCommand(sql, sqlCon, sqlTrn);
        //        sqlCmd.ExecuteNonQuery();
        //    }
        //    catch (Exception e)
        //    {
        //        throw;
        //    }
        //}

        ////以下ほぼ自動生成コード
        ////protected override void dispose(Boolean disposing)
        ////{
        ////    if(!(disposedValue))
        ////    {
        ////        Me.Close();
        ////        disposedValue = true;
        ////    }
        ////}

        ////protected override void finalize()
        ////{

        ////}

        ////public void dispose() 
    }
}
