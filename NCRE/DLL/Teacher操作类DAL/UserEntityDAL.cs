using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Model;

namespace DAL
{
    /// <summary>
    /// 针对用户表进行操作类--2015年11月14日11:50:11--周洲
    /// </summary>
    public class UserEntityDAL
    {
        private SQLHelper sqlhelper = null;
        public UserEntityDAL()
        {
            sqlhelper =new SQLHelper(); 
        }
        public DataTable  TeacherLoginByName(UserEntity userinfo)
        {
        
                DataTable dt = new DataTable();
                string sql = "select * from UserEntity where userName=@userName and userPassword =@PWD ";
                SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@userName",userinfo.userName),
                new SqlParameter ("@PWD",userinfo.userPassword) };
                dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
                return dt;
           
            
        }
    }
}
