using DAL;
using Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;

namespace BLL
{
    /// <summary>
    /// 考试配置逻辑类 （含基础数据的读取）
    /// </summary>
    public class ExamConfigBLL
    {
        #region 共有属性
        SelectQuestionEntityDAL selectQuestionDal = new SelectQuestionEntityDAL();
        SelectQuestionRecordEntityDAL selectQuestionRecordDal = new SelectQuestionRecordEntityDAL();
        CollegeDAL collegeDal = new CollegeDAL();

        DynamicCreationDAL dynamicCreateDAL = new DynamicCreationDAL();
        #endregion

        #region 清空 单个学生的  选择题答题记录
        /// <summary>
        /// 清空 单个学生的  选择题答题记录
        /// </summary>
        /// <param name="enStudent"></param>
        /// <returns></returns>
        public int ClearSelectQuestionRecordByStudent(StudentInfoEntity enStudent)
        {
            return selectQuestionRecordDal.ClearSelectQuestionRecordByStudent(enStudent);
        } 
        #endregion

        #region 清空 学生List  的选择题答题记录
        /// <summary>
        /// 清空 学生List  的选择题答题记录
        /// </summary>
        /// <param name="lstStudent">学生列表</param>
        /// <returns>返回0表示删除成功，返回-1表示删除失败</returns>
        public int ClearSelectQuestionRecordByLstStudent(List<StudentInfoEntity> lstStudent)
        {
            return selectQuestionRecordDal.ClearSelectQuestionRecordByLstStudent(lstStudent);
        } 
        #endregion

        #region 清空某指定学院内的所有选择题答题记录
        /// <summary>
        /// 清空某指定学院内的所有选择题答题记录
        /// </summary>
        /// <param name="pEnCollege">学院实体</param>
        /// <returns>返回0表示删除成功，返回-1表示删除失败</returns>
        public int DeleteSelectQuestionRecordByCollegeID(CollegeEntity pEnCollege)
        {
            string tableName = "SelectQuestionRecordEntity_" + pEnCollege.collegeID;
            List<string> lstTableName=new List<string>();
            lstTableName.Add(tableName);

            //方法一：真清空
            return selectQuestionRecordDal.ClearSelectQuestionRecordByCollegeID(pEnCollege);
            
            ////方法二：假删除  学生的答题记录表
            //if (true==dynamicCreateDAL.FalseDropTable(null, lstTableName, null))
            //{
            //    return 0;
            //}
            ////假删除失败
            //return -1;
        }

        /// <summary>
        /// 删除 指定 学生Lst的 选择题答题记录
        /// </summary>
        /// <param name="pLstStudent">学生列表</param>
        /// <returns></returns>
        public int DeleteSelectQuestionRecordByLstStudent(List<StudentInfoEntity> pLstStudent) {
            return ClearSelectQuestionRecordByLstStudent(pLstStudent);
        }

        /// <summary>
        /// 清空某指定学院内的所有选择题答题记录
        /// </summary>
        /// <param name="pEnCollege">学院实体</param>
        /// <returns>返回0表示删除成功，返回-1表示删除失败</returns>
        public int FalseClearSelectQuestionRecordByCollegeID(CollegeEntity pEnCollege)
        {
            string tableName = "SelectQuestionRecordEntity_" + pEnCollege.collegeID;
            List<string> lstTableName = new List<string>();
            lstTableName.Add(tableName);

            //方法二：假删除  学生的答题记录表
            if (true == dynamicCreateDAL.FalseDropTable(null, lstTableName, null))
            {
                return 0;
            }
            //假删除失败
            return -1;
        }
        #endregion

        #region 清空所有选择题答题记录
        /// <summary>
        /// 清空  ALL   选择题答题记录 By 学院Lst
        /// </summary>
        /// <param name="pEnCollege">学院实体</param>
        /// <returns>返回0表示删除成功，返回-1表示删除失败</returns>
        public int ClearSelectQuestionRecordByCollegeID(List<CollegeEntity> pLstCollege)
        {
            return selectQuestionRecordDal.ClearSelectQuestionRecordByCollegeID(pLstCollege);
        }
        #endregion

        #region 获取所有的学院
        /// <summary>
        /// 获取所有的学院
        /// </summary>
        /// <returns></returns>
        public List<CollegeEntity> GetAllCollege()
        {
            return collegeDal.GetAllCollege();
        }
        #endregion

        #region 根据学院获取学生
        /// <summary>
        /// 根据学院获取学生
        /// </summary>
        /// <param name="pCollege"></param>
        /// <returns></returns>
        public List<StudentInfoEntity> GetStudentByCollege(CollegeEntity pCollege)
        {
            return collegeDal.GetStudentByCollege(pCollege);
        }
        #endregion

        #region 对该学生列表 随机抽题，生成答题记录
        /// <summary>
        /// 对该学生列表 随机抽题，生成答题记录
        /// </summary>
        /// <param name="lstStudent">学生列表</param>
        public void RandGenerateRecord(List<StudentInfoEntity> lstStudent)
        {
            //获取所有的选择题
            List<SelectQuestionEntity> allSelectQuestion = selectQuestionDal.LoadAllSelectQuestion();
            //随机得到的选择题
            List<SelectQuestionEntity> selectedQuestion = null;
            int QuestionCount = allSelectQuestion.Count();

            string strSelectQuestionCount =ConfigurationManager.ConnectionStrings["selectQuestionCount"].ConnectionString;
            int count = int.Parse(strSelectQuestionCount);

            //1，对每个学生进行抽题操作
            foreach (StudentInfoEntity enStudent in lstStudent)
            {
                List<int> listRandom = GetListRandom(count, QuestionCount);

                selectedQuestion = new List<SelectQuestionEntity>();
                selectedQuestion = GetSomeQuestion(listRandom, allSelectQuestion);
                //将随机生成的学生答题记录保存到相应的记录表中

                //要创建的答题记录表集合
                string tableName = "";
                tableName = "SelectQuestionRecordEntity_" + enStudent.CollegeID ;
                List<string> lstTableName = new List<string>();
                lstTableName.Add(tableName); 

                if (!dynamicCreateDAL.IsTableExist(null,tableName,null))
                {
                    //如果不存在，则创建该学院的答题记录表  复制选择题答题记录模版
                    dynamicCreateDAL.CreateDataTableCopySelectRecord(null, lstTableName, null, "SelectQuestionRecordEntity");
                }

                #region 向答题记录表中插入随机的答题记录
                //执行效率太慢。100个学生     1分钟
                //selectQuestionRecordDal.InsertStudentSelectQuestionRecord(enStudent, selectedQuestion);
                //执行效率      1000个学生     12s  不过数据类型转换有问题
                //selectQuestionRecordDal.InsertTableRecord(enStudent, selectedQuestion);
                //执行效率      1000个学生     20s
                //DateTime dt = DateTime.Now;
                //for (int i = 0; i < 1000; i++)
                //{
                    selectQuestionRecordDal.InsertRecordByUnion(enStudent, selectedQuestion); 
                //}
                //TimeSpan ts = DateTime.Now - dt;
                #endregion

            }
        }
        #endregion

        #region 个不重复随机数GetListRandom
        /// <summary>
        /// 取count个不重复随机数
        /// </summary>
        /// <param name="count">取的个数</param>
        /// <param name="Max_Count">最大范围</param>
        /// <returns>随机数的list集合</returns>
        private List<int> GetListRandom(int count, int Max_Count)
        {
            Random random = new Random();
            int intRan = 0;
            List<int> listRandom = new List<int>();   //定义一个集合，用来存储生成的随机数

            //生成count个随机不相同的数
            for (int i = 0; i < count; i++)
            {
                intRan = Convert.ToInt32(random.Next(1, Max_Count));

                if (listRandom.Contains(intRan))
                {
                    i--;
                }
                else
                {
                    listRandom.Add(intRan);
                }

            }

            return listRandom;    //利用三目运算确定区间的开始位置
        }
        #endregion

        #region 取listRandom.count个数的实体集合
        /// <summary>
        /// 取listRandom.count个数的实体集合
        /// </summary>
        /// <param name="listRandom"></param>
        /// <param name="listTotalQuestion"></param>
        /// <returns></returns>
        public List<SelectQuestionEntity> GetSomeQuestion(List<int> listRandom, List<SelectQuestionEntity> listTotalQuestion)
        {
            List<SelectQuestionEntity> listQuestionTemp = new List<SelectQuestionEntity>();
            //for (int i = 0; i < listTotalQuestion.Count; i++)
            //{
            //    for (int j = 0; j < listRandom.Count; j++)
            //    {
            //        if (listTotalQuestion[i].QuestionID.Trim() == listRandom[j].ToString().Trim())
            //        {
            //            int id = j + 1;
            //            listTotalQuestion[i].QuestionID = id.ToString();
            //            listQuestionTemp.Add(listTotalQuestion[i]);
            //        }
            //    }
            //}

            for (int i = 0; i < listRandom.Count; i++)
            {
                SelectQuestionEntity temp = listTotalQuestion.Find(s => s.QuestionID.ToString().Trim() == listRandom[i].ToString().Trim());
                int num =i+1;
                temp.QuestionID = listRandom[i].ToString().Trim();
                listQuestionTemp.Add(temp);
            }

            return listQuestionTemp;
        }
        #endregion


        #region NCRE 方法
        /// <summary>
        /// NCRE:判断数据库表是否存在
        /// </summary>
        /// <param name="tb">数据表名，必须指定</param>
        /// <returns>true:表示数据表已经存在；false，表示数据表不存在</returns>
        public Boolean IsTableExist(string tb)
        {
            return dynamicCreateDAL.IsTableExist(null, tb, null);
        }

        /// <summary>
        /// NCRE:复制答题记录表
        /// </summary>
        /// <param name="strDataTable">要创建的数据表</param>
        /// <param name="WantCopyTable">要复制的表名</param>
        public void CreateDataTableCopySelectRecord(string strDataTable, string WantCopyTable)
        {
            List<string> lstTableName = new List<string>();
            lstTableName.Add(strDataTable);
            dynamicCreateDAL.CreateDataTableCopySelectRecord(null, lstTableName, null, WantCopyTable);
        } 
        #endregion
    }
}
