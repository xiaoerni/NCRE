using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using Model;

namespace DAL
{
    public class StudentScoreDAL
    {
        private SQLHelper sqlhelper =null;
        public StudentScoreDAL (){
            sqlhelper =new SQLHelper ();
        }

        #region 从学生表，学院表，成绩表中联合查出学生的成绩并显示到前面---周洲---2015年11月14日20:28:42

        /// <summary>
        /// 从学生表，学院表，成绩表中联合查出学生的成绩并显示到前面---周洲---2015年11月14日20:28:42
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectStudentByMajor(StudentInfoEntity studentinfo){
            DataTable dt=new DataTable ();
            string sql="select studentID as 学号 ,studentName as 姓名,(select collegeName  from CollegeEntity where collegeID =StudentInfoEntity.collegeID)as 学院,major as 专业,(Select score from ScoreEntity where studentID =StudentInfoEntity.studentID )as 得分 from StudentInfoEntity where major=@major";
            SqlParameter[] paras = new SqlParameter[] {
                new SqlParameter ("@major",studentinfo.major)};
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;

        }
        #endregion


        #region 从学生表，学院表，成绩表中联合查出学生的成绩并显示到前面---周洲---2015年11月14日20:28:42

        /// <summary>
        /// 从学生表，学院表，成绩表中联合查出学生的成绩并显示到前面---周洲---2015年11月14日20:28:42
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectScoreByCollege(StudentInfoEntity studentinfo)
        {
            DataTable dt = new DataTable();
            string sql = "select studentID as 学号 ,studentName as 姓名,(select collegeName  from CollegeEntity where collegeID =StudentInfoEntity.collegeID)as 学院,major as 专业,(Select score from ScoreEntity where studentID =StudentInfoEntity.studentID )as 得分 from StudentInfoEntity where collegeID=@collegeID";
            SqlParameter[] paras = new SqlParameter[] {
                new SqlParameter ("@collegeID",studentinfo.CollegeID )};
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;

        }
        #endregion


        #region 选择所有需要考试学生的专业---周洲--2015年11月16日
        /// <summary>
        /// 选择所有需要考试学生的专业---周洲--2015年11月16日
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectAllMajor() {
            DataTable dt = new DataTable();
            string sql = "select distinct major from StudentInfoEntity";
            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);
            return dt;
        }
        #endregion

        #region 选择所有需要考试学生的学院---周洲--2015年11月16日
        /// <summary>
        /// 选择所有需要考试学生的专业---周洲--2015年11月16日
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectAllCollege()
        {
            DataTable dt = new DataTable();
            string sql = "select collegeName from CollegeEntity";
            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);
            return dt;
        }
        #endregion

        #region 选择所有需要考试学生的学院ID和姓名---周洲--2015年11月16日
        /// <summary>
        /// 选择所有需要考试学生的学院ID和姓名---周洲--2015年11月16日
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectAllCollegeInfo()
        {
            DataTable dt = new DataTable();
            string sql = "select * from CollegeEntity";
            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);
            return dt;
        }
        #endregion

        #region 选择对应学院下拉框的专业---周洲--2015年11月17日
        /// <summary>
        /// 选择对应学院下拉框的专业---周洲--2015年11月17日
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectMajorByCollegeID(StudentInfoEntity studentinfo)
        {
        
            DataTable dt = new DataTable();
            string sql = "select distinct major from StudentInfoEntity where collegeID=@collegeID";
            SqlParameter[] paras = new SqlParameter[] {
                new SqlParameter ("@collegeID",studentinfo.CollegeID )};
            dt = sqlhelper.ExecuteQuery(sql,paras,CommandType.Text);
            return dt;
        }
        #endregion


    }

}
