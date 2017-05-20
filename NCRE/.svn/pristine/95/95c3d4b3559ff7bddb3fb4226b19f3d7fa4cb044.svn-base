/********************************************************************* 
作者：赵崇
小组：开发小组(十期考评系统：一清、刘晓春、赵崇、霍亚静、刘杰、刘新阳、马巧盼、任焱、杨晓敏、连江伟、孟海滨、肖红、王潇峥）
说明：
创建日期：2015/4/2 17:22:52
版本号： V1.0.0
**********************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Data;
using DAL;



namespace DAL
{
    public class DynamicCreationDAL
    {
        #region NCRE 数据：所用的数据库名db  和数据库连接connKey
        private string NCREdb = ConfigurationManager.ConnectionStrings["NCREdb"].ConnectionString;
        private string NCREconnKey = "connStr";
        #endregion

        #region 判断数据库是否存在
        /// <summary>
        /// 判断数据库是否存在
        /// </summary>
        /// <param name="db">数据库的名称，可以传入null 默认是NCRE的数据库</param>
        /// <param name="connKey">数据库的连接Key，可以传入null，默认是使用NCRE的connKey</param>
        /// <returns>true:表示数据库已经存在；false，表示数据库不存在</returns>
        public Boolean IsDBExist(string db, string connKey)
        {
            SQLHelper helper = new SQLHelper();

            string connToMaster = ConfigurationManager.ConnectionStrings[connKey == null ? NCREconnKey : connKey].ToString();
            string tableName = (db == null ? NCREdb : db);
            string createDbStr = " select * from master.dbo.sysdatabases where name " + "= '" + tableName + "'";

            DataTable dt = helper.ExecuteQuery(createDbStr, CommandType.Text);
            if (dt.Rows.Count == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        #endregion

        #region 判断数据库中，指定表是否存在
        /// <summary>
        /// 判断数据库表是否存在
        /// </summary>
        /// <param name="db">数据库，可以传入null 默认是NCRE的数据库</param>
        /// <param name="tb">数据表名，必须指定</param>
        /// <param name="connKey">连接数据库的key，可以传入null，默认是使用NCRE的connKey</param>
        /// <returns>true:表示数据表已经存在；false，表示数据表不存在</returns>
        public Boolean IsTableExist(string db, string tb, string connKey)
        {
            SQLHelper helper = new SQLHelper();

            string connToMaster = ConfigurationManager.ConnectionStrings[connKey == null ? NCREconnKey : connKey].ToString();
            string strDB = (db == null ? NCREdb : db);
            string createDbStr = "use \"" + strDB + "\" select 1 from  sysobjects where  id = object_id('" + tb + "') and type = 'U'";

            //在指定的数据库中  查找 该表是否存在
            DataTable dt = helper.ExecuteQuery(createDbStr, CommandType.Text);
            if (dt.Rows.Count == 0)
            {
                return false;
            }
            else
            {
                return true;
            }

        }
        #endregion

        #region 创建数据库 赵崇-2014年12月24日 15:53:08
        /// <summary>
        /// 创建数据库
        /// </summary>
        /// <param name="db">数据库名称，可以传入null 默认是NCRE的数据库</param>
        /// <param name="connKey">连接数据库的key，可以传入null，默认是使用NCRE的connKey</param>
        /// <returns>返回true 数据库创建成功</returns>
        public bool CreateDataBase(string db, string connKey)
        {
            SQLHelper helper = new SQLHelper();
            //符号变量，判断数据库是否存在
            Boolean flag = IsDBExist(db, connKey);

            //如果数据库存在，则抛出
            if (flag == true)
            {
                return false;
            }
            else
            {
                //数据库不存在，创建数据库
                string connToMaster = ConfigurationManager.ConnectionStrings[connKey == null ? NCREconnKey : connKey].ToString();
                string strDB=(db == null ? NCREdb : db);
                string createDbStr = "Create database " + strDB;
                helper.ExecuteNonQuery(createDbStr, CommandType.Text);
                return true;
            }

        }
        #endregion

        #region 删除数据库 赵崇-2014年12月24日 15:51:48
        /// <summary>
        /// 删除数据库
        /// </summary>
        /// <param name="db">数据库名，可以传入null 默认是NCRE的数据库</param>
        /// <param name="connKey">数据库连接串，可以传入null，默认是使用NCRE的connKey</param>
        /// <returns>删除成功为true，删除失败为false</returns>
        public bool DropDataBase(string db, string connKey)
        {
            SQLHelper helper = new SQLHelper();
            //符号变量，判断数据库是否存在
            Boolean flag = IsDBExist(db, connKey);

            //如果数据库不存在，则抛出
            if (flag == false)
            {
                return false;
            }
            else
            {
                //数据库存在，删除数据库
                string connToMaster = ConfigurationManager.ConnectionStrings[connKey == null ? NCREconnKey : connKey].ToString();
                string strDB = (db == null ? NCREdb : db);
                string createDbStr = "Drop database " + strDB;
                helper.ExecuteNonQuery(createDbStr, CommandType.Text);
                return true;
            }
        }
        #endregion

        #region 创建数据库表(多张表)
        /// <summary>
        ///  在指定的数据库中，创建数据表
        /// </summary>
        /// <param name="db">指定的数据库，可以传入null 默认是NCRE的数据库</param>
        /// <param name="listDt">要创建的数据表集合</param>
        /// <param name="dic">数据表中的字段及其数据类型  Dictionary集合</param>
        /// <param name="connKey">数据库的连接Key，可以传入null，默认是使用NCRE的connKey</param>
        public void CreateDataTable(string db, List<string> listDt, Dictionary<string, string> dic, string connKey)
        {
            SQLHelper helper = new SQLHelper();

            string connToMaster = ConfigurationManager.ConnectionStrings[connKey == null ? NCREconnKey : connKey].ToString();

            //判断数据库是否存在
            if (IsDBExist(db, connKey) == false)
            {
                throw new Exception("数据库不存在！");
            }

            for (int i = 0; i < listDt.Count(); i++)
            {
                //如果数据库表存在，则抛出错误
                if (IsTableExist(db, listDt[i], connKey) == true)
                {
                    //如果数据库表已经存在，则跳过该表
                    continue;
                }
                else//数据表不存在，创建数据表
                {
                    //其后判断数据表是否存在，然后创建数据表
                    string strDB = (db == null ? NCREdb : db);
                    string createTableStr = GenerateExecuteSql(strDB, listDt[i], dic);
                    helper.ExecuteNonQuery(createTableStr, CommandType.Text);
                }
            }
        }
        #endregion

        #region 批量删除数据表
        /// <summary>
        /// 批量删除数据库
        /// </summary>
        /// <param name="db">指定的数据库，可以传入null 默认是NCRE的数据库</param>
        /// <param name="listDt">要删除的数据库表集合</param>
        /// <param name="connKey">数据库连接串，可以传入null，默认是使用NCRE的connKey</param>
        /// <returns>删除是否成功，true表示删除成功，false表示删除失败</returns>
        public bool DropDataTable(string db, List<string> listDt, string connKey)
        {
            SQLHelper helper = new SQLHelper();

            string connToMaster = ConfigurationManager.ConnectionStrings[connKey == null ? NCREconnKey : connKey].ToString();

            //判断数据库是否存在
            if (IsDBExist(db, connKey) == false)
            {
                throw new Exception("数据库不存在！");
            }

            for (int i = 0; i < listDt.Count(); i++)
            {
                //如果数据库表存在，则抛出错误
                if (IsTableExist(db, listDt[i], connKey) == false)
                {
                    //如果数据库表已经删除，则跳过该表
                    continue;
                }
                else//数据表存在，则进行删除数据表
                {
                    //其后判断数据表是否存在，然后创建数据表
                    string strDB=  (db == null ? NCREdb : db);
                    string createTableStr = "use \"" + strDB + "\" drop table " + listDt[i] + " ";
                    helper.ExecuteNonQuery(createTableStr, CommandType.Text);
                }
            }
            return true;
        }
        #endregion

        #region 批量假删除数据表（重命名，打上时间戳)  赵崇-2015年2月4日 10:10:56
        /// <summary>
        /// 批量假删除数据表（重命名，打上时间戳）
        /// </summary>
        /// <param name="db">指定的数据库，可以传入null 默认是NCRE的数据库</param>
        /// <param name="listDt">要删除的数据库表集合</param>
        /// <param name="connKey">数据库连接串，可以传入null，默认是使用NCRE的connKey</param>
        /// <returns>执行是否成功，true表示成功，false表示失败</returns>
        public bool FalseDropTable(string db, List<string> listDt, string connKey)
        {
            SQLHelper helper = new SQLHelper();

            string connToMaster = ConfigurationManager.ConnectionStrings[connKey == null ? NCREconnKey : connKey].ToString();

            //判断数据库是否存在
            if (IsDBExist(db, connKey) == false)
            {
                throw new Exception("数据库不存在！");
            }

            string timeStamp = DateTime.Now.ToString();            // 2008-9-4 20:02:10

            for (int i = 0; i < listDt.Count(); i++)
            {
                //如果数据库表存在，则抛出错误
                if (IsTableExist(db, listDt[i], connKey) == false)
                {
                    //如果数据库表已经删除，则跳过该表
                    continue;
                }
                else//数据表存在，则进行删除数据表
                {
                    //其后判断数据表是否存在，然后创建数据表
                    //string createTableStr = "use " + db + " drop table " + listDt[i] + " ";
                    string strDB = (db == null ? NCREdb : db);
                    string createTableStr = "use \"" + strDB + "\" exec sp_rename '" + listDt[i] + "','" + listDt[i] + "_" + timeStamp + "'";
                    helper.ExecuteNonQuery(createTableStr, CommandType.Text);
                }
            }
            return true;
        }
        #endregion

        #region 拼接创建数据库表的Sql语句 赵崇-2014年12月24日 15:09:19
        /// <summary>
        /// 拼接创建数据库表的Sql语句
        /// </summary>
        /// <param name="db">指定的数据库</param>
        /// <param name="dt">要创建的数据表</param>
        /// <param name="dic">数据表中的字段及其数据类型</param>
        /// <returns>拼接完的字符串</returns>
        public string GenerateExecuteSql(string db, string dt, Dictionary<string, string> dic)
        {
            //拼接字符串，（该串为创建内容）
            string content = "serial int identity(1,1) primary key ";
            //取出dic中的内容，进行拼接
            List<string> test = new List<string>(dic.Keys);
            for (int i = 0; i < dic.Count(); i++)
            {
                content = content + " , " + test[i] + " " + dic[test[i]];
            }

            //其后判断数据表是否存在，然后创建数据表
            string strDB = (db == null ? NCREdb : db);
            string createTableStr = "use \"" + strDB + "\" create table " + dt + " (" + content + ")";
            return createTableStr;
        }
        #endregion


        #region 创建答题记录表(多张表)  复制模版
        /// <summary>
        ///  在指定的数据库中，创建数据表
        /// </summary>
        /// <param name="db">指定的数据库，可以传入null 默认是NCRE的数据库</param>
        /// <param name="listDt">要创建的数据表集合</param>
        /// <param name="connKey">数据库的连接Key，可以传入null，默认是使用NCRE的connKey</param>
        /// <param name="WantCopyTable">要复制的表名</param>
        public void CreateDataTableCopySelectRecord(string db, List<string> listDt, string connKey,string WantCopyTable)
        {
            SQLHelper helper = new SQLHelper();

            string connToMaster = ConfigurationManager.ConnectionStrings[connKey == null ? NCREconnKey : connKey].ToString();

            //判断数据库是否存在
            if (IsDBExist(db, connKey) == false)
            {
                throw new Exception("数据库不存在！");
            }

            for (int i = 0; i < listDt.Count(); i++)
            {
                //如果数据库表存在，则抛出错误
                if (IsTableExist(db == null ? NCREdb : db, listDt[i], connKey == null ? NCREconnKey : connKey) == true)
                {
                    //如果数据库表已经存在，则跳过该表
                    continue;
                }
                else//数据表不存在，创建数据表
                {
                    //其后判断数据表是否存在，然后  复制 第一张答题记录表
                    string createTableStr = "use " + db == null ? NCREdb : db + " select * into " + listDt[i] + "  from " + WantCopyTable + " where 1=0";
                    helper.ExecuteNonQuery(createTableStr, CommandType.Text);
                }
            }
        }
        #endregion
    }
}
