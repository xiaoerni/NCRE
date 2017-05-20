using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using Model;

namespace DAL
{
    public class ExcelEntityDAL
    {
        private SQLHelper sqlhelper = null;

        public ExcelEntityDAL()
        {
            sqlhelper = new SQLHelper();
        }

        #region 根据考试试题类型ID判分——王虹芸
        /// <summary>
        /// 根据考试试题类型ID进行判分
        /// </summary>
        /// <param name="excelinfo">试题信息excelinfo</param>
        /// <returns>查询结果dt</returns>
        public int UpdateExcelTypeID(ExcelQuestionRecordEntity excelrecord)
        {
            String which = WhichExcelRecored(excelrecord);
            DataTable dt = new DataTable();
            string sql = "update ExcelQuestionRecordEntity_" + which + " set Fration=0 ,ExamAnswer='考生未答题' where PaperType=@PaperType";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter ("@PaperType",excelrecord.PaperType )
            };
            int flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);

            return flag;
        }
        #endregion

        #region 根据考试试题类型ID查询试卷——王虹芸
        /// <summary>
        /// 根据考试试题类型ID进行判分
        /// </summary>
        /// <param name="excelinfo">试题信息excelinfo</param>
        /// <returns>查询结果dt</returns>
        public DataTable QueryExcelTypeID(ExcelQuestionEntity excelinfo)
        {
            DataTable dt = new DataTable();
            string sql = "select * from ExcelQuestionEntity where QuestionFlag=@QuestionFlag and PaperType=@PaperType";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter ("@QuestionFlag",excelinfo.QuestionFlag ),
                new SqlParameter ("@PaperType",excelinfo.PaperType )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);

            return dt;
        }
        #endregion

        #region 根据papertype从题库表中获取试题信息--周洲--2015年11月21日
        /// <summary>
        /// 加载excel试题信息到界面中
        /// </summary>
        /// <param name="excelinfo">学号</param>
        /// <returns>试题信息</returns>
        public DataTable LoadExcelQuestion(ExcelQuestionEntity excelinfo)
        {
            DataTable dt = new DataTable();
            string sql = "select * from ExcelQuestionEntity where PaperType =@PaperType";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter ("@PaperType",excelinfo.PaperType)
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion

        #region excel判分之后返回信息到数据库——王虹芸
        /// <summary>
        /// 判分之后返回信息到数据库
        /// </summary>
        /// <param name="excelrecord">根据StudentID判断</param>
        /// <returns>是否添加成功</returns>
        public int ReturnExcelScore(ExcelQuestionRecordEntity excelrecord)
        {
            String which = WhichExcelRecored(excelrecord);
            string sql = "update ExcelQuestionRecordEntity_" + which + " set ExamAnswer=@ExamAnswer,Fration=@Fration where QuestionID=@QuestionID and PaperType=@PaperType";
            SqlParameter[] paras = new SqlParameter[] {
                new SqlParameter ("@QuestionID",excelrecord.QuestionID ),
                new SqlParameter ("@PaperType",excelrecord.PaperType ),
                new SqlParameter ("@Fration",excelrecord.Fration ),
                new SqlParameter ("@ExamAnswer",excelrecord.ExamAnswer )              
            };
            int flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            return flag;
        }
        #endregion

        #region excel查询试题类型关键字（QuestionType）——王虹芸
        /// <summary>
        /// 查询试题类型关键字
        /// </summary>
        /// <param name="excelrecord">根据QuestionTypeID判断</param>
        /// <returns></returns>
        public DataTable QueryQuestionType(ExcelQuestionEntity exceltype)
        {
            DataTable dt = new DataTable();
            string sql = "select QuestionFlag from ExcelQuestionEntity where PaperType=@PaperType";
            SqlParameter[] paras = new SqlParameter[] {
                  new SqlParameter ("@PaperType",exceltype.PaperType )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
            return dt;
        }
        #endregion

        #region 查找所有的Excel的套卷 ExcelPaperType() 邱慕夏 2015年11月20日16:57:30
        /// <summary>
        /// 查找所有的Excel的套卷
        /// </summary>
        /// <returns></returns>
        public DataTable ExcelPaperType()
        {
            DataTable dt = new DataTable();
            string sql = "select Distinct PaperType from ExcelQuestionEntity";
            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);
            return dt;
        }
        #endregion

        #region 给ExcelQuestionEntity分组 邱慕夏 2015年11月20日16:57:30
        /// <summary>
        /// 查找所有的Excel的套卷
        /// </summary>
        /// <returns></returns>
        public DataTable ExcelPaperTypeGroupByPaperType()
        {
            DataTable dt = new DataTable();
            string sql = "select * from ExcelQuestionEntity group by PaperType";
            dt = sqlhelper.ExecuteQuery(sql, CommandType.Text);
            return dt;
        }
        #endregion


        #region 根据PaperType查询分组 邱慕夏 2015年11月20日16:57:30
        /// <summary>
        ///给ExcelQutionEntity分组
        /// </summary>
        /// <returns></returns>
        public DataTable ExcelPaperTypeGroupByPaperType(String PaperType)
        {
            DataTable dt = new DataTable();
            string sql = "select * from ExcelQuestionEntity where PaperType=@PaperType";
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
        public int InsertExcelRecord(ExcelQuestionRecordEntity studentrecord)
        {
            String which = WhichExcelRecored(studentrecord);
            string sql = "Insert into ExcelQuestionRecordEntity_" + which + "(QuestionID,StudentID,PaperType,QuestionContent,CorrectAnswer) values(@QuestionID,@StudentID,@PaperType,@QuestionContent,@RightAnswer)";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecord.StudentID ),
                new SqlParameter ("@PaperType",studentrecord.PaperType ),
                new SqlParameter ("@QuestionContent",studentrecord.QuestionContent),
                new SqlParameter ("@RightAnswer",studentrecord.CorrectAnswer),
                new SqlParameter ("@QuestionID",studentrecord.QuestionID )
            };
            int flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            return flag;
        }
        #endregion

        #region 根据学生的ID查询去重--邱慕夏  2015年11月23日14:28:57
        /// <summary>
        /// 根据学生的ID查询去重--邱慕夏--邱慕夏
        /// </summary>
        /// <param name="studentinfo">根据CollegeID判断学生在那个表中</param>
        /// <returns>是否已经添加</returns>
        public Boolean SelectExcelRecord(ExcelQuestionRecordEntity studentrecord)
        {
            String which = WhichExcelRecored(studentrecord);
            string sql = "Select * from  ExcelQuestionRecordEntity_" + which + " where StudentID=@studentID";
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
        public int InsertExcelRecordList(List<ExcelQuestionRecordEntity> studentrecordlist)
        {
            int flag = 0;
            for (int i = 0; i < studentrecordlist.Count; i++)
            {
                String which = WhichExcelRecored(studentrecordlist[i]);
                string sql = "Insert into ExcelQuestionRecordEntity_" + which + "(QuestionID,StudentID,PaperType,QuestionContent,CorrectAnswer) values(@QuestionID,@StudentID,@PaperType,@QuestionContent,@RightAnswer)";
                SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@studentID",studentrecordlist[i].StudentID ),
                new SqlParameter ("@PaperType",studentrecordlist[i].PaperType ),
                new SqlParameter ("@QuestionContent",studentrecordlist[i].QuestionContent),
                new SqlParameter ("@RightAnswer",studentrecordlist[i].CorrectAnswer),
                new SqlParameter ("@QuestionID",studentrecordlist[i].QuestionID )
                };
                flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            }
            return flag;
        }
        #endregion

        #region 看看向哪个数据库表中进行填写数据 WhichExcelRecored(ExcelQuestionRecordEntity studentrecord)  邱慕夏  2015年11月23日14:14:49
        /// <summary>
        /// 看看向哪个数据库表中进行填写数据
        /// </summary>
        /// <param name="studentrecord"></param>
        /// <returns></returns>
        public String WhichExcelRecored(ExcelQuestionRecordEntity studentrecord)
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
            tableName = "ExcelQuestionRecordEntity_" + pEnStudent.CollegeID;

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
