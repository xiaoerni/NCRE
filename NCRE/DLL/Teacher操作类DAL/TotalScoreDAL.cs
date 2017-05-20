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
    /// 得到每一个题的得分情况 李少然
    /// </summary>
    public class TotalScoreDAL
    {
        private SQLHelper sqlhelper = null;

        public TotalScoreDAL()
        {
            sqlhelper = new SQLHelper();
        }
        #region 调出word分数,修改人李少然
        //word分数
        public DataTable WordTotalScore(StudentInfoEntity studentinfo)
        {
            String which = WhichWordRecored(studentinfo);
            DataTable dt = new DataTable();
            string sql = "select StudentID as '学号',ExamAnswer as '学生答案',RightAnswer as '正确答案',Fration as '分数' from WordQuestionRecordEntity_"+which +" where StudentID =@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@studentID",studentinfo.studentID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion

        #region 调出ppt分数 修改人：李少然
        //ppt分数
        public DataTable PptTotalScore(StudentInfoEntity studentinfo)
        {
            String which = WhichWordRecored(studentinfo);
            DataTable dt = new DataTable();
            string sql = "select StudentID as '学号',ExamAnswer as '学生答案',RightAnswer as '正确答案',Fration as '分数' from PptQuestionRecordEntity_"+which +"   where StudentID =@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@studentID",studentinfo.studentID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion

        #region 调出Windows分数
        public DataTable WindowsScore(StudentInfoEntity studentinfo)
        {
            String which = WhichWordRecored(studentinfo);
            DataTable dt = new DataTable();
            string sql = "select StudentID as '学号',examAnswer as '学生答案',correctAnswer as '正确答案',fraction as '分数'from WinQuestionRecordEntity_"+which +" where studentID =@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@studentID",studentinfo.studentID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion

        #region 调出IE分数
        //ie分数
        public DataTable IETotalScore(StudentInfoEntity studentinfo)
        {
            String which = WhichWordRecored(studentinfo);
            DataTable dt = new DataTable();
            string sql = " select studentID as '学号',examAnswer as '学生答案',correctAnswer as '正确答案',fraction as '分数' from IEQuestionRecordEntity_"+which +" where studentID =@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@studentID",studentinfo.studentID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion

        #region 调出Excel分数
        //excel分数
        public DataTable ExcelTotalScore(StudentInfoEntity studentinfo)
        {
            String which = WhichWordRecored(studentinfo);
            DataTable dt = new DataTable();
            string sql = "select StudentID as '学号',ExamAnswer as '学生答案',CorrectAnswer as '正确答案',Fration as '分数' from ExcelQuestionRecordEntity_"+which +" where studentID =@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@studentID",studentinfo.studentID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion

        #region 看看向哪个数据库表中进行填写数据  邱慕夏  2015年11月23日14:26:41
        /// <summary>
        /// 看看向哪个数据库表中进行填写数据
        /// </summary>
        /// <param name="studentrecord"></param>
        /// <returns></returns>
        public String WhichWordRecored(StudentInfoEntity studentrecord)
        {
            DataTable dt = new DataTable();
            string sql = "select collegeId  from StudentInfoEntity where studentID =@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecord.studentID  ),
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            String which = dt.Rows[0][0].ToString();
            return which;
        }
        #endregion


    }
}

