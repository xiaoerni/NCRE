/*
 * 创建人：Cindy
 * 创建时间：2015年6月1日16:08:35
 * 说明：数据库的助手类
 * 版权所有：TGB韩梦甜
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace DAL
{
    public class SQLHelper
    {
        private SqlConnection conn = null;
        private SqlCommand cmd = null;
        private SqlDataReader sdr = null;

        //private SqlConnection GetConnection(string strDatabaseName)
        //{
        //    string strConnValue = ConfigHelper.ReadAppSetting(strDatabaseName);
        //    return new SqlConnection(strConnValue);
        //}

        #region"连接数据库"
        /// <summary>
        /// 连接数据库
        /// </summary>
        public SQLHelper()
        {
            string connStr = ConfigurationManager.ConnectionStrings["connStr"].ConnectionString;
            conn = new SqlConnection(connStr);
        }
        #endregion
        #region "打开数据库连接"
        /// <summary>
        /// 打开数据库连接
        /// </summary>
        /// <returns></returns>
        public SqlConnection GetConn()
        {
            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
            return conn;
        }
        #endregion
        #region "该方法执行不带参数的增删改SQL语句或存储过程"
        /// <summary>
        /// 该方法执行不带参数的增删改SQL语句或存储过程
        /// </summary>
        /// <param name="cmdText">要执行的SQL语句或存储 过程</param>
        ///<param name ="ct">命令类型</param>
        /// <returns>返回更新的记录数</returns>
        public int ExecuteNonQuery(string cmdText, CommandType ct)
        {
            int res;
            try
            {
                cmd = new SqlCommand(cmdText, GetConn());
                cmd.CommandType = ct;
                res = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
            }
            return res;
        }
        #endregion
        #region“该方法执行带参数的增删改SQL语句”
        /// <summary>
        /// 该方法执行带参数的增删改SQL语句
        /// </summary>
        /// <param name="cmdText"> /param>
        /// <returns></returns>
        public int ExecuteNonQuery(string cmdText, SqlParameter[] paras, CommandType ct)
        {
            int res;
            using (cmd = new SqlCommand(cmdText, GetConn()))
            {
                cmd.CommandType = ct;
                cmd.Parameters.AddRange(paras);
                res = cmd.ExecuteNonQuery();
            }
            return res;
        }
        #endregion
        #region "该方法执行传入的SQL查询语句"
        /// <summary>
        /// 该方法执行传入的SQL查询语句或存储过程
        /// </summary>
        /// <param name="cmdText">SQL查询语句或存储过程</param>
        /// <returns></returns>
        public DataTable ExecuteQuery(string cmdText, CommandType ct)
        {
            DataTable dt = new DataTable();
            cmd = new SqlCommand(cmdText, GetConn());
            cmd.CommandType = ct;
            using (sdr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
            {
                dt.Load(sdr);
            }
            return dt;
        }
        #endregion
        #region "该方法执行带参数的SQL查询语句"
        /// <summary>
        /// 该方法执行带参数的SQL查询语句
        /// </summary>
        /// <param name="cmdText">要执行的SQL语句</param>
        /// <param name="paras">参数集合</param>
        /// <returns></returns>
        public DataTable ExecuteQuery(string cmdText, SqlParameter[] paras, CommandType ct)
        {
            DataTable dt;

            dt = new DataTable();
            cmd = new SqlCommand(cmdText, GetConn());
            cmd.CommandType = ct;
            cmd.Parameters.AddRange(paras);
            using (sdr = cmd.ExecuteReader(CommandBehavior.CloseConnection))
            {
                dt.Load(sdr);
            }
            return dt;
        }
        #endregion

        public bool ExecuteNonQuery(List<string> sqlList)
        {
            //using (SQLiteConnection con = new SQLiteConnection(connStr))
            //{
            //    con.Open();
            //    DbTransaction trans = con.BeginTransaction();//开始事务       
            //    SQLiteCommand cmd = new SQLiteCommand(con);
            //    try
            //    {
            //        cmd.CommandText = "INSERT INTO MyTable(username,useraddr,userage) VALUES(@a,@b,@c)";
            //        for (int n = 0; n < 100000; n++)
            //        {
            //            cmd.Parameters.Add(new SQLiteParameter("@a", DbType.String)); //MySql 使用MySqlDbType.String  
            //            cmd.Parameters.Add(new SQLiteParameter("@b", DbType.String)); //MySql 引用MySql.Data.dll  
            //            cmd.Parameters.Add(new SQLiteParameter("@c", DbType.String));
            //            cmd.Parameters["@a"].Value = "张三" + n;
            //            cmd.Parameters["@b"].Value = "深圳" + n;
            //            cmd.Parameters["@c"].Value = 10 + n;
            //            cmd.ExecuteNonQuery();
            //        }
            //        trans.Commit();//提交事务    
            //        DateTime endtime = DateTime.Now;
            //        MessageBox.Show("插入成功，用时" + (endtime - starttime).TotalMilliseconds);

            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message);
            //    }
            //}  
            return true;
        }

        #region 批量导入DataTable

        public void BulkToDB(string tableName, DataTable dt)
        {
            SqlConnection sqlConn = new SqlConnection(
                ConfigurationManager.ConnectionStrings["connStr"].ConnectionString);
            SqlBulkCopy bulkCopy = new SqlBulkCopy(sqlConn);
            bulkCopy.DestinationTableName = tableName;
            bulkCopy.BatchSize = dt.Rows.Count;

            try
            {
                sqlConn.Open();
                if (dt != null && dt.Rows.Count != 0)
                    bulkCopy.WriteToServer(dt);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                sqlConn.Close();
                if (bulkCopy != null)
                    bulkCopy.Close();
            }
        }

        #endregion 批量导入DataTable
    }
}
