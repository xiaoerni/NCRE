
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient ;
using Model;
using DAL;

namespace DAL
{
    public class IEQuestionEntityDAL
    {
        private SQLHelper sqlhelper;

        #region "实例化SQLHelper"
        /// <summary>
        /// SQLHelper实例化
        /// </summary>
        public IEQuestionEntityDAL()
        {
            sqlhelper = new SQLHelper();
        }
        #endregion

        #region  根据题型选择出IE的题干--周洲--2015年11月21日
        /// <summary>
        /// 根据题型选择出IE的题干--周洲--2015年11月21日
        /// </summary>
        /// <param name="iequestion">传递考试类型</param>
        /// <returns>返回题目内容</returns>
        public DataTable LoadIEQuestion(IEQuestionEntity ieinfo)
        {

            DataTable ieQuestionDt = new DataTable();
            string cmdText = "select * from IEQuestionEntity where paperType =@paperType";
            SqlParameter[] paras = new SqlParameter[]{new SqlParameter ("@paperType",ieinfo.paperType )
        };
            ieQuestionDt = sqlhelper.ExecuteQuery(cmdText, paras, CommandType.Text);
            return ieQuestionDt;

        }
        #endregion


        #region "加载试题信息到界面中，以试卷类型判断，查出题型-韩梦甜-2015-11-20"
        /// <summary>
        /// 加载试题信息到界面中，以试卷类型判断，查出题型
        /// </summary>
        /// <param name="iequestion">传递考试类型</param>
        /// <returns>返回题目内容</returns>
        public DataTable LoadIEByFlag(IEQuestionEntity iequestion)
        {
            DataTable ieQuestionDt = new DataTable();
            string cmdText = "select questionID,questionContent,questionFlag,fraction,paperType,correctAnswer from IEQuestionEntity where paperType=@paperType and questionFlag=@questionFlag";

            SqlParameter[] paras = new SqlParameter[]{new SqlParameter ("@paperType",iequestion.paperType) ,
                new  SqlParameter("@questionFlag",iequestion.questionFlag )
        };
            ieQuestionDt = sqlhelper.ExecuteQuery(cmdText, paras, CommandType.Text);
            return ieQuestionDt;

        }
        #endregion

        #region "根据学号选择题型-韩梦甜-2015-11-20"
        /// <summary>
        /// 根据学号选择题型-韩梦甜-2015-11-20
        /// </summary>
        /// <param name="studentinfo">传递学号</param>
        /// <returns>返回题目内容</returns>
        public DataTable SelectPaperTypeByStudentID(StudentInfoEntity studentinfo)
        {
            DataTable Dt = new DataTable();
            string cmdText = "select paperType from IEQuestionRecordEntity where studentID=@studentinfo";

            SqlParameter[] paras = new SqlParameter[]{new SqlParameter ("@studentinfo",studentinfo.studentID )};
            Dt = sqlhelper.ExecuteQuery(cmdText, paras, CommandType.Text);
            return Dt;

        }
        #endregion

        #region"判分之后，将学生的答案传到答题记录表-韩梦甜-2015-11-20"
        /// <summary>
        /// 判分之后，将学生的答案传到答题记录表
        /// </summary>
        /// <param name="iequestion">根据考生ID判断</param>
        /// <returns>是否更新成功</returns>
        public int ReturnScore(IEQuestionRecordEntity studentRecord)
        {
            String which = WhichIERecored(studentRecord);
            string cmdText = "Update  IEQuestionRecordEntity_" + which + " set  examAnswer=@examAnswer,  fraction=@fraction where questionID=@questionID and studentID=@studentID ";
            SqlParameter[] paras = new SqlParameter[]{new SqlParameter ("@studentID",studentRecord.studentID ),
                new SqlParameter ("@questionID",studentRecord.questionID ),
                new SqlParameter ("@fraction",studentRecord.fraction ),
                new SqlParameter ("@examAnswer",studentRecord.examAnswer )};

           int flag= sqlhelper.ExecuteNonQuery(cmdText, paras, CommandType.Text);
           return flag;
        }
        #endregion

        #region 查找所有的IE的套卷 IEPaperType() 邱慕夏 2015年11月20日16:57:30
        /// <summary>
        /// 查找所有的word的套卷
        /// </summary>
        /// <returns></returns>
        public DataTable IEPaperType()
        {
            DataTable dt = new DataTable();
            string sql = "select Distinct paperType from IEQuestionEntity";
            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);
            return dt;
        }
        #endregion

        #region 给WordQutionEntity分组 邱慕夏 2015年11月20日16:57:30
        /// <summary>
        /// 查找所有的word的套卷
        /// </summary>
        /// <returns></returns>
        public DataTable IEPaperTypeGroupByPaperType()
        {
            DataTable dt = new DataTable();
            string sql = "select * from IEQuestionEntity group by paperType";
            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);
            return dt;
        }
        #endregion


        #region 根据PaperType查询分组 邱慕夏 2015年11月20日16:57:30
        /// <summary>
        ///给WordQutionEntity分组
        /// </summary>
        /// <returns></returns>
        public DataTable IEPaperTypeGroupByPaperType(String PaperType)
        {
            DataTable dt = new DataTable();
            string sql = "select * from IEQuestionEntity where paperType=@PaperType";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@PaperType",PaperType )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion


        #region 根据学生的ID查询是否该学生是要往哪个表中进行insert--邱慕夏
        /// <summary>
        /// 根据学生的ID查询是否该学生是要往哪个表中进行insert--邱慕夏
        /// </summary>
        /// <param name="studentinfo">根据CollegeID判断学生在那个表中</param>
        /// <returns>是否添加成功</returns>
        public int InsertIERecord(IEQuestionRecordEntity studentrecord)
        {
            String which = WhichIERecored(studentrecord);
            string sql = "Insert into IEQuestionRecordEntity_" + which + "(questionID,studentID,paperType,questionContent,correctAnswer) values(@QuestionID,@StudentID,@PaperType,@QuestionContent,@RightAnswer)";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecord.studentID ),
                new SqlParameter ("@PaperType",studentrecord.paperType ),
                new SqlParameter ("@QuestionContent",studentrecord.questionContent),
                new SqlParameter ("@RightAnswer",studentrecord.correctAnswer),
                new SqlParameter ("@QuestionID",studentrecord.questionID )
            };
            int flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            return flag;
        }
        #endregion

        #region 根据学生的ID查询去重--邱慕夏  2015年11月23日14:28:57
        /// <summary>
        /// 根据学生的ID查询去重--邱慕夏
        /// </summary>
        /// <param name="studentinfo">根据CollegeID判断学生在那个表中</param>
        /// <returns>是否已经添加</returns>
        public Boolean SelectIERecord(IEQuestionRecordEntity studentrecord)
        {
            String which = WhichIERecored(studentrecord);
            string sql = "Select * from  IEQuestionRecordEntity_" + which + " where studentID=@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecord.studentID ),
            };
            DataTable dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            if (dt.Rows.Count == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region 批量根据学生的ID查询是否该学生是要往哪个表中进行insert--邱慕夏
        /// <summary>
        /// 批量根据学生的ID查询是否该学生是要往哪个表中进行insert--邱慕夏
        /// </summary>
        /// <param name="studentinfo">根据CollegeID判断学生在那个表中</param>
        /// <returns>是否添加成功</returns>
        public int InsertIERecordList(List<IEQuestionRecordEntity> studentrecordlist)
        {
            int flag = 0;
            for (int i = 0; i < studentrecordlist.Count; i++)
            {
                String which = WhichIERecored(studentrecordlist[i]);
                string sql = "Insert into IEQuestionRecordEntity_" + which + "(questionID,studentID,paperType,questionContent,correctAnswer) values(@QuestionID,@StudentID,@PaperType,@QuestionContent,@RightAnswer)";
                SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter ("@QuestionID",studentrecordlist[i].questionID ),
                new SqlParameter("@StudentID",studentrecordlist[i].studentID ),
                new SqlParameter ("@PaperType",studentrecordlist[i].paperType ),
                new SqlParameter ("@QuestionContent",studentrecordlist[i].questionContent),
                new SqlParameter ("@RightAnswer",studentrecordlist[i].correctAnswer)
                
                };
                flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            }
            return flag;
        }
        #endregion

        #region 看看向哪个数据库表中进行填写数据--邱慕夏 -2015年11月23日14:17:36
        /// <summary>
        /// 看看向哪个数据库表中进行填写数据
        /// </summary>
        /// <param name="studentrecord"></param>
        /// <returns></returns>
        public String WhichIERecored(IEQuestionRecordEntity studentrecord)
        {
            DataTable dt = new DataTable();
            string sql = "select collegeId from StudentInfoEntity where studentID =@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecord.studentID),
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            String which = dt.Rows[0][0].ToString();
            return which;
        }
        #endregion

        #region 查询该考生 是否有资格考试 -赵崇-2015年11月24日 16:39:57
        /// <summary>
        /// 查询该考生 是否有资格考试
        /// </summary>
        /// <param name="pEnStudent">学生</param>
        /// <returns>true表示可以进行考试，false表示 未进行配置 所以不能考试</returns>
        public bool GetIsCanExamByStudent(StudentInfoEntity pEnStudent)
        {

            string tableName = "";
            tableName = "IEQuestionRecordEntity_" + pEnStudent.CollegeID;

            //1,查询所有的选择题
            DataTable dt = new DataTable();
            string sql = "select * from " + tableName + " where StudentID=@StudentID";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter ("@StudentID",pEnStudent.studentID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);

            if (dt.Rows.Count == 0)
            {
                return false;
            }
            return true;
        }
        #endregion
    }
}
