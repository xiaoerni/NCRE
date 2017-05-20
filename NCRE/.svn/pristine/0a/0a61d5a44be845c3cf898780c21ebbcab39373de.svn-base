using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;
using System.Data;
using System.Data.SqlClient;

namespace DAL
{
    public class PptQuestionEntityDAL
    {
        private SQLHelper sqlhelper = null;
        public  PptQuestionEntityDAL()
        {
            sqlhelper = new SQLHelper();
        }
        /// <summary>
        /// 加载试题信息到界面
        /// </summary>
        /// <param name="pptinfo">传递学生号</param>
        /// <returns>返回题目的内容</returns>
        public DataTable LoadPptByFlag(PptQuestionEntity pptinfo)
        {
            DataTable dt = new DataTable();
            string sql = "select QuestionID,QuestionFlag,QuestionContent,Fration, PaperType,RightAnswer from  PptQuestionEntity  where PaperType=@PaperType and QuestionFlag=@QuestionFlag";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@PaperType",pptinfo.PaperType  ),
                new SqlParameter ("@QuestionFlag",pptinfo.QuestionFlag )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }

        #region 根据试题的类型加载试题的信息--周洲--2015年11月21日
        /// <summary>
        /// 加载试题到界面中
        /// </summary>
        /// <param name="pptinfo">传递学生号</param>
        /// <returns>返回题目的内容</returns>
        public DataTable LoadPptQuestion(PptQuestionEntity pptinfo)
        {
            DataTable dt = new DataTable();
            string sql = "select * from PptQuestionEntity where paperType =@PaperType";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@PaperType",pptinfo.PaperType )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion

        /// <summary>
        /// 判分之后返回信息到数据库
        /// </summary>
        /// <param name="studentinfo">根据考生ID判断</param>
        /// <returns>是否添加成功</returns>
        public int ReturnScord(PptQuestionRecordEntity studentrecord)
        {
            String which = WhichPptRecored(studentrecord);
            string sql = "update PptQuestionRecordEntity_" + which + "   set Fration=@Fration,ExamAnswer =@ExamAnswer where QuestionID =@QuestionID and studentID =@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecord.StudentID ),
                new SqlParameter ("@Fration",studentrecord.Fration ),
                new SqlParameter ("@ExamAnswer",studentrecord.ExamAnswer),
                new SqlParameter ("@QuestionID",studentrecord.QuestionID )
            };
            int flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            return flag;
        }

        #region 根据学号选择题型  李少然
        /// <summary>
        /// 根据学号选择题型
        /// </summary>
        /// <param name="studentinfo">传递学生号</param>
        /// <returns>试卷类型</returns>
        public DataTable SelectPaperTypeByStudentID(StudentInfoEntity studentinfo)
        {
            DataTable dt = new DataTable();
            string sql = "select PaperType from PptQuestionRecordEntity   where StudentID = @studentinfo";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@studentinfo",studentinfo.studentID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;

        }
        #endregion

        #region 查找所有的Ppt的套卷 PptPaperType() 邱慕夏 2015年11月20日16:57:30
        /// <summary>
        /// 查找所有的Ppt的套卷
        /// </summary>
        /// <returns></returns>
        public DataTable PptPaperType()
        {
            DataTable dt = new DataTable();
            string sql = "select Distinct PaperType from PptQuestionEntity";
            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);
            return dt;
        }
        #endregion

        #region 给PptQutionEntity分组 邱慕夏 2015年11月20日16:57:30
        /// <summary>
        /// 查找所有的Ppt的套卷
        /// </summary>
        /// <returns></returns>
        public DataTable PptPaperTypeGroupByPaperType()
        {
            DataTable dt = new DataTable();
            string sql = "select * from PptQuestionEntity group by PaperType";
            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);
            return dt;
        }
        #endregion


        #region 根据PaperType查询分组 邱慕夏 2015年11月20日16:57:30
        /// <summary>
        ///给PptQutionEntity分组
        /// </summary>
        /// <returns></returns>
        public DataTable PptPaperTypeGroupByPaperType(String PaperType)
        {
            DataTable dt = new DataTable();
            string sql = "select * from PptQuestionEntity where PaperType=@PaperType";
            SqlParameter[] paras = new SqlParameter[]{
                new  SqlParameter("@PaperType",PaperType )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion


        #region 根据学生的ID查询是否该学生是要往哪个表中进行insert--邱慕夏  2015年11月23日14:28:57
        /// <summary>
        /// 根据学生的ID查询是否该学生是要往哪个表中进行insert--邱慕夏
        /// </summary>
        /// <param name="studentinfo">根据CollegeID判断学生在那个表中</param>
        /// <returns>是否添加成功</returns>
        public int InsertPptRecord(PptQuestionRecordEntity studentrecord)
        {
            String which = WhichPptRecored(studentrecord);
            string sql = "Insert into PptQuestionRecordEntity_" + which + "(QuestionID,StudentID,PaperType,QuestionContent,RightAnswer) values(@QuestionID,@StudentID,@PaperType,@QuestionContent,@RightAnswer)";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecord.StudentID ),
                new SqlParameter ("@PaperType",studentrecord.PaperType ),
                new SqlParameter ("@QuestionContent",studentrecord.QuestionContent),
                new SqlParameter ("@RightAnswer",studentrecord.RightAnswer),
                new SqlParameter ("@QuestionID",studentrecord.QuestionID )
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
        public Boolean SelectPptRecord(PptQuestionRecordEntity studentrecord)
        {
            String which = WhichPptRecored(studentrecord);
            string sql = "Select * from  PptQuestionRecordEntity_" + which + " where StudentID=@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecord.StudentID ),
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
        public int InsertPptRecordList(List<PptQuestionRecordEntity> studentrecordlist)
        {
            int flag = 0;
            for (int i = 0; i < studentrecordlist.Count; i++)
            {
                String which = WhichPptRecored(studentrecordlist[i]);
                string sql = "Insert into PptQuestionRecordEntity_" + which + "(QuestionID,StudentID,PaperType,QuestionContent,RightAnswer) values(@QuestionID,@StudentID,@PaperType,@QuestionContent,@RightAnswer)";
                SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecordlist[i].StudentID ),
                new SqlParameter ("@PaperType",studentrecordlist[i].PaperType ),
                new SqlParameter ("@QuestionContent",studentrecordlist[i].QuestionContent),
                new SqlParameter ("@RightAnswer",studentrecordlist[i].RightAnswer),
                new SqlParameter ("@QuestionID",studentrecordlist[i].QuestionID )
                };
                flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            }
            return flag;
        }
        #endregion

        #region 看看向哪个数据库表中进行填写数据  邱慕夏  2015年11月23日14:26:41
        /// <summary>
        /// 看看向哪个数据库表中进行填写数据
        /// </summary>
        /// <param name="studentrecord"></param>
        /// <returns></returns>
        public String WhichPptRecored(PptQuestionRecordEntity studentrecord)
        {
            DataTable dt = new DataTable();
            string sql = "select collegeId  from StudentInfoEntity where studentID =@studentID";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecord.StudentID ),
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
            tableName = "PptQuestionRecordEntity_" + pEnStudent.CollegeID;

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
