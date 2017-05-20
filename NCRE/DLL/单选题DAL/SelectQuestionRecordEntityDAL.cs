using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Model;
using System.Data.SqlClient;

namespace DAL
{
    public class SelectQuestionRecordEntityDAL
    {
        private SQLHelper sqlhelper = null;

        public SelectQuestionRecordEntityDAL()
        {
            sqlhelper = new SQLHelper();
        }

        TBToList<SelectQuestionRecordEntity> dtToList = new TBToList<SelectQuestionRecordEntity>();

        #region 查询该考生 是否有资格考试 -赵崇-2015年11月24日
        /// <summary>
        /// 查询该考生 是否有资格考试
        /// </summary>
        /// <param name="pEnStudent">学生</param>
        /// <returns>true表示可以进行考试，false表示 未进行配置 所以不能考试</returns>
        public bool GetIsCanExamByStudent(StudentInfoEntity pEnStudent)
        {

            string tableName = "";
            tableName = "SelectQuestionRecordEntity_" + pEnStudent.CollegeID ;

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

        #region 根据 考生ID 和 学院ID 查询该考生 所抽到的试题-赵崇-2015年11月16日 19:29:05
        /// <summary>
        /// 根据 考生ID 和 学院ID 查询该考生 所抽到的试题
        /// </summary>
        /// <param name="pStudentId"></param>
        /// <param name="pCollege"></param>
        /// <returns></returns>
        public List<SelectQuestionRecordEntity> GetLstSelectQuestionRecordByStudentIdAndCollegeId(StudentInfoEntity pEnStudent)
        {
            string tableName = "";
            tableName = "SelectQuestionRecordEntity_" + pEnStudent.CollegeID;

            //1,查询所有的选择题
            DataTable dt = new DataTable();
            string sql = "select * from " + tableName + " where StudentID=@StudentID";
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter ("@StudentID",pEnStudent.studentID )
            };
            dt = sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);

            //2,把查询结果组织成  实体List
            List<SelectQuestionRecordEntity> lsSelectQuestionRecordEntity = new List<SelectQuestionRecordEntity>();
            lsSelectQuestionRecordEntity = dtToList.ToList(dt);

            return lsSelectQuestionRecordEntity;
        }
        #endregion

        #region 保存考生的答题记录信息（Single） -赵崇- 2015年11月16日 19:29:23
        /// <summary>
        /// 保存考生的答题记录信息（Single）
        /// </summary>
        /// <param name="pEnStudentInfo">学生信息</param>
        /// <param name="pEnSelectRecord">答题记录信息</param>
        /// <returns></returns>
        public int UpdateSelectQuestionRecordByStudentInfo(StudentInfoEntity pEnStudentInfo, SelectQuestionRecordEntity pEnSelectRecord,string rightAnswer)
        {
            string tableName = "";
            string sql = string.Empty;
            tableName = "SelectQuestionRecordEntity_" + pEnStudentInfo.CollegeID;

            //1，根据考生信息 更新相应的答题记录
                //判断学生答题是否正确
            if (rightAnswer != string.Empty)
            {

                if (rightAnswer == pEnSelectRecord.ExamAnswer)
                {
                    sql = "update " + tableName + " set ExamAnswer=@ExamAnswer,Fration = '2' where StudentID=@StudentID  and QuestionID=@QuestionID ";
                }
                else
                {
                    sql = "update " + tableName + " set ExamAnswer=@ExamAnswer,Fration = '0' where StudentID=@StudentID  and QuestionID=@QuestionID";
                }

            }
            else 
            { 
                sql = "update " + tableName + " set ExamAnswer=@ExamAnswer where StudentID=@StudentID  and QuestionID=@QuestionID";
            }
           
            SqlParameter[] paras = new SqlParameter[]{
                new SqlParameter("@StudentID",pEnStudentInfo.studentID ),
                new SqlParameter ("@ExamAnswer", pEnSelectRecord.ExamAnswer),
                new SqlParameter ("@QuestionID", pEnSelectRecord.QuestionID)
            };
            //2,返回保存结果
            int flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            return flag;
        }
        #endregion


        public void UpdateFrationInfo(StudentInfoEntity pEnStudentInfo, SelectQuestionRecordEntity pEnSelectRecord)
        {
            string tableName = "";
            tableName = "SelectQuestionRecordEntity_" + pEnStudentInfo.CollegeID;
        
        }

        #region 未实现   更新考生的答题记录信息(List) 赵崇-2015年11月16日 19:29:41
        ///// <summary>
        ///// 更新考生的答题记录信息(List)
        ///// </summary>
        ///// <param name="pEnStudentInfo">学生信息</param>
        ///// <param name="pLstSelectRecord">答题记录集合</param>
        ///// <returns></returns>
        //public int UpdateSelectQuestionRecordByStudentInfo(StudentInfoEntity pEnStudentInfo, List<SelectQuestionRecordEntity  > pLstSelectRecord)
        //{
        //    //1，根据考生信息 更新相应的答题记录
        //    //string sql = "update SelectQuestionRecordEntity_@CollegeID set RightAnswer=@RightAnswer where studentID=@studentID and QuestionID=@QuestionID";
        //    //SqlParameter[] paras = new SqlParameter[]{
        //    //    new SqlParameter("@studentID",pEnStudentInfo.studentID ),
        //    //    new SqlParameter ("@collegeID",pEnStudentInfo.collegeID),
        //    //    new SqlParameter ("@RightAnswer", pLstSelectRecord[0].RightAnswer),
        //    //    new SqlParameter ("@QuestionID", pLstSelectRecord[0].QuestionID)
        //    //};

        //    List<string> lstSql = new List<string>();
        //    string sql = "";

        //    foreach (SelectQuestionRecordEntity   record in pLstSelectRecord)
        //    {
        //        sql = string.Format("update SelectQuestionRecordEntity_{0} set RightAnswer='{1}' where QuestionID='{2}'", pEnStudentInfo.collegeID, record.RightAnswer, record.QuestionID);
        //        lstSql.Add(sql);
        //    }

        //    //2,返回保存结果
        //    //int flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);

        //    return 0;
        //}
        #endregion

        #region 清空

        #region 清空 单个 学生的 选择题答题记录
        /// <summary>
        /// 清空 单个 学生的 选择题答题记录
        /// </summary>
        /// <param name="enStudent"></param>
        /// <returns></returns>
        public int ClearSelectQuestionRecordByStudent(StudentInfoEntity enStudent)
        {
            try
            {
                string tableName = "SelectQuestionRecordEntity_" + enStudent.CollegeID;
                string sql = "delete from " + tableName + " where studentID=@studentID";
                SqlParameter[] paras = new SqlParameter[]{
                    new SqlParameter("@studentID",enStudent.studentID )
                };
                sqlhelper.ExecuteQuery(sql, paras, CommandType.Text);
                return 0;
            }
            catch (Exception)
            {
                return -1;
            }
        }
        #endregion

        #region 清空  学生List 的选择题答题记录
        /// <summary>
        /// 清空  学生List 的选择题答题记录
        /// </summary>
        /// <param name="lstStudent"></param>
        /// <returns></returns>
        public int ClearSelectQuestionRecordByLstStudent(List<StudentInfoEntity> lstStudent)
        {
            try
            {
                foreach (StudentInfoEntity student in lstStudent)
                {
                    ClearSelectQuestionRecordByStudent(student);
                }
                return 0;
            }
            catch (Exception)
            {
                return -1;
            }
        }  
        #endregion

        #region 清空某指定学院内的所有选择题答题记录  by college -赵崇-2015年11月16日 23:33:42
        /// <summary>
        /// 清空某指定学院内的所有选择题答题记录
        /// </summary>
        /// <param name="pEnCollege">学院实体</param>
        /// <returns>返回0表示删除成功，返回-1表示删除失败</returns>
        public int ClearSelectQuestionRecordByCollegeID(CollegeEntity pEnCollege)
        {

            try
            {
                string tableName = "SelectQuestionRecordEntity_" + pEnCollege.collegeID;
                string sql = "delete from " + tableName + " ";
                sqlhelper.ExecuteQuery(sql, CommandType.Text);
                return 0;
            }
            catch (Exception)
            {
                return -1;
            }
        }
        #endregion

        #region 清空所有选择题答题记录  -赵崇-2015年11月16日 23:33:42
        /// <summary>
        /// 清空所有选择题答题记录
        /// </summary>
        /// <param name="pEnCollege">学院实体</param>
        /// <returns>返回0表示删除成功，返回-1表示删除失败</returns>
        public int ClearSelectQuestionRecordByCollegeID(List<CollegeEntity> pLstCollege)
        {
            try
            {
                foreach (CollegeEntity college in pLstCollege)
                {
                    ClearSelectQuestionRecordByCollegeID(college);
                }
                return 0;
            }
            catch (Exception)
            {
                return -1;
            }
        }
        #endregion
        #endregion

        #region 抽题
        #region 将随机生成的学生答题记录保存到相应的记录表中  低效率
        /// <summary>
        /// 将随机生成的学生答题记录保存到相应的记录表中
        /// </summary>
        /// <param name="lstQuestion"></param>
        /// <returns></returns>
        public int InsertStudentSelectQuestionRecord(StudentInfoEntity pEnStudentInfo, List<SelectQuestionEntity> lstQuestion)
        {
            string tableName = "";
            tableName = "SelectQuestionRecordEntity_" + pEnStudentInfo.CollegeID;

            //1，根据考生信息 更新相应的答题记录
            foreach (SelectQuestionEntity item in lstQuestion)
            {
                string sql = "insert into " + tableName + "(QuestionID,StudentID,PaperType,QuestionContent,AnswerA,AnswerB,AnswerC,AnswerD,RightAnswer,Fration) values(@QuestionID,@StudentID,@PaperType,@QuestionContent,@AnswerA,@AnswerB,@AnswerC,@AnswerD,@RightAnswer,@Fration)";
                SqlParameter[] paras = new SqlParameter[]{
                    new SqlParameter("@QuestionID", item.QuestionID),
                    new SqlParameter("@StudentID", pEnStudentInfo.studentID),
                    new SqlParameter("@PaperType", "-"),
                    new SqlParameter("@QuestionContent", item.QuestionContent),
                    new SqlParameter("@AnswerA",item.OptionA ),
                    new SqlParameter("@AnswerB", item.OptionB),
                    new SqlParameter("@AnswerC", item.OptionC),
                    new SqlParameter("@AnswerD",item.OptionD ),
                    new SqlParameter("@RightAnswer",item.RightAnswer ),
                    new SqlParameter("@Fration", item.Fration),
                };

                //2,返回保存结果
                int flag = sqlhelper.ExecuteNonQuery(sql, paras, CommandType.Text);
            }
            return 0;
        }
        #endregion

        #region 将随机生成的学生答题记录保存到相应的记录表中  dataTable
        /// <summary>
        /// 将随机生成的学生答题记录保存到相应的记录表中
        /// </summary>
        /// <param name="lstQuestion"></param>
        /// <returns></returns>
        public int InsertTableRecord(StudentInfoEntity pEnStudentInfo, List<SelectQuestionEntity> lstQuestion)
        {
            string tableName = "";
            tableName = "SelectQuestionRecordEntity_" + pEnStudentInfo.CollegeID;

            DataTable dt = new DataTable();
            dt = dtToList.ToDataTableTow(lstQuestion);
            //Console.WriteLine(dt.Rows.Count);
            sqlhelper.BulkToDB(tableName, dt);
            //2,返回保存结果
            return 0;
        }
        #endregion

        #region 随机生成的学生答题记录保存到相应的记录表中 Union方式
        /// <summary>
        /// 随机生成的学生答题记录保存到相应的记录表中 Union方式
        /// </summary>
        /// <param name="enStudent"></param>
        /// <param name="selectedQuestion"></param>
        public void InsertRecordByUnion(StudentInfoEntity pEnStudentInfo, List<SelectQuestionEntity> lstQuestion)
        {
            #region 1，生成随机的选择题
            string tableName = "";
            tableName = "SelectQuestionRecordEntity_" + pEnStudentInfo.CollegeID;

            StringBuilder sbSql = new StringBuilder();

            sbSql.Append("insert into ").Append(tableName.ToString()).Append(" (QuestionID,StudentID,PaperType,QuestionContent,OptionA,OptionB,OptionC,OptionD,RightAnswer)");

            //foreach (SelectQuestionEntity item in lstQuestion)
            //{
            //    sbSql.Append(" select '" + item.QuestionID + "',")
            //        .Append("'" + pEnStudentInfo.studentID + "',")
            //        .Append("'" + "-----" + "',")
            //        .Append("'" + item.QuestionContent + "',")
            //        .Append("'" + item.OptionA + "',")
            //        .Append("'" + item.OptionB + "',")
            //        .Append("'" + item.OptionC + "',")
            //        .Append("'" + item.OptionD + "',")
            //        .Append("'" + item.RightAnswer + "',")
            //        .Append("'" + item.Fration + "' " + " union all ");
            //}

            for (int i = 0; i < lstQuestion.Count -1; i++)
            {
                sbSql.Append(" select '" + lstQuestion[i].QuestionID + "',")
               .Append("'" + pEnStudentInfo.studentID + "',")
               .Append("'" + "-----" + "',")
               .Append("'" + lstQuestion[i].QuestionContent + "',")
               .Append("'" + lstQuestion[i].OptionA + "',")
               .Append("'" + lstQuestion[i].OptionB + "',")
               .Append("'" + lstQuestion[i].OptionC + "',")
               .Append("'" + lstQuestion[i].OptionD + "',")
               .Append("'" + lstQuestion[i].RightAnswer + "'" + " union all ");
            }

            sbSql.Append(" select '" + lstQuestion[lstQuestion.Count - 1].QuestionID + "',")
              .Append("'" + pEnStudentInfo.studentID + "',")
              .Append("'" + "-----" + "',")
              .Append("'" + lstQuestion[lstQuestion.Count - 1].QuestionContent + "',")
              .Append("'" + lstQuestion[lstQuestion.Count - 1].OptionA + "',")
              .Append("'" + lstQuestion[lstQuestion.Count - 1].OptionB + "',")
              .Append("'" + lstQuestion[lstQuestion.Count - 1].OptionC + "',")
              .Append("'" + lstQuestion[lstQuestion.Count - 1].OptionD + "',")
              .Append("'" + lstQuestion[lstQuestion.Count - 1].Fration + "'");

            //去除最后一个Union all
            //sbSql.Remove(sbSql.Length - 12, 12);
            #endregion

            //2，执行sql语句，返回保存结果
            int flag = sqlhelper.ExecuteNonQuery(sbSql.ToString(), CommandType.Text);
        }
        #endregion
        #endregion
    }
}
