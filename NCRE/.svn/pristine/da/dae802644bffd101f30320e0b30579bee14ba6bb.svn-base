using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Model;

namespace DAL
{
    public class StudentInfoEntityDAL
    {
        private SQLHelper sqlhelper = null;

        public StudentInfoEntityDAL()
        {
            sqlhelper =new SQLHelper(); 
        }

        TBToList<StudentInfoEntity> dtToList = new TBToList<StudentInfoEntity>();

        #region 选择学生信息通过学号---周洲--2015-11-21


        /// <summary>
        /// 选择学生信息通过学号---周洲--2015-11-21
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectStudentInfoByID(StudentInfoEntity studentinfo)
        {
            DataTable dt = new DataTable();
            string sql = "select * from StudentInfoEntity where studentID=@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@studentID",studentinfo.studentID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion

        #region 通过studentID选择学生的试卷类型（A,B,C）--周洲--2015年11月21日
        /// <summary>
        /// 通过studentID选择学生的试卷类型（A,B,C）--周洲--2015年11月21日
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectPaperTypebyStudentID(StudentInfoEntity studentinfo)
        {
            DataTable dt = new DataTable();
            string sql = "select * from StudentBindPaperTypeEntity where StudentID=@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@studentID",studentinfo.studentID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;

        }
        #endregion


        #region 通过学号 查询学生信息 -赵崇-2015年11月24日 17:01:58
        /// <summary>
        /// 通过学号 查询学生信息
        /// </summary>
        /// <param name="studentId">学号</param>
        /// <returns>学生信息</returns>
        public StudentInfoEntity GetStudentById(string studentId)
        {
            DataTable dt = new DataTable();
            string sql = "select * from StudentInfoEntity where studentID=@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@studentID",studentId )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);

            //2,把查询结果组织成  实体List
            List<StudentInfoEntity> lstStudent = new List<StudentInfoEntity>();
            lstStudent = dtToList.ToList(dt);

            return lstStudent[0];
        } 
        #endregion
    }
}
