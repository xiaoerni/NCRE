using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DAL;
using Model;
using System.Data;
using System.Data.SqlClient;

namespace BLL
{
    public class ExcelEntityBLL
    {
        private ExcelEntityDAL exceldal = null;
        public ExcelEntityBLL()
        {
            exceldal = new ExcelEntityDAL();
        }

        /**
       * qmx
       * **/
        private TBToList<ExcelQuestionEntity> tBToList = new TBToList<ExcelQuestionEntity>();


        public DataTable QueryExcelTypeID(ExcelQuestionEntity excel)
        {
            return exceldal.QueryExcelTypeID(excel);
        }

        public int UpdateExcelTypeID(ExcelQuestionRecordEntity excelrecord)
        {
            return exceldal.UpdateExcelTypeID(excelrecord);
        }

        public int UpdateExcelRecord(ExcelQuestionRecordEntity excelrecord)
        {
            // return exceldal.ReturnExcelScore (excelrecord);
            int flag;
            flag = exceldal.ReturnExcelScore(excelrecord);
            return flag;
        }

        #region 根据papertype从题库表中获取试题信息--周洲--2015年11月21日
        /// <summary>
        /// 根据papertype从题库表中获取试题信息--周洲--2015年11月21日
        /// </summary>
        /// <param name="excelinfo"></param>
        /// <returns></returns>
        public DataTable LoadExcelQuestion(ExcelQuestionEntity excelinfo)
        {
            return exceldal.LoadExcelQuestion(excelinfo);
        }
        #endregion

        #region 返回excel分值到数据库当中——王虹芸
        /// <summary>
        /// 返回分值到数据库当中
        /// </summary>
        /// <param name="excelrecord">正确答案</param>
        /// <returns>返回学号，得分，学生答案，正确答案</returns>
        public int ReturnExcelScore(ExcelQuestionRecordEntity excelrecord)
        {
            int flag = exceldal.ReturnExcelScore(excelrecord);
            return flag;
        }
        #endregion

        #region excel查询试题类型关键字（QuestionType）——王虹芸
        /// <summary>
        /// excel查询试题类型关键字（QuestionType）
        /// </summary>
        /// <param name="excelinfo"></param>
        /// <returns></returns>
        public DataTable QueryExcelQuestionType(ExcelQuestionEntity exceltype)
        {
            return exceldal.QueryQuestionType(exceltype);
        }
        #endregion

        #region 查找所有的Excel的套卷 ExcelPaperType() 邱慕夏 2015年11月20日16:57:30
        public DataTable ExcelPaperType()
        {
            return exceldal.ExcelPaperType();
        }
        #endregion

        #region 给ExcelQutionEntity分组 邱慕夏 2015年11月20日16:57:30
        public List<ExcelQuestionEntity> ExcelPaperTypeGroupByPaperType(String PaperType)
        {

            return tBToList.ToList(exceldal.ExcelPaperTypeGroupByPaperType(PaperType));
        }
        #endregion

        #region 根据学生的ID查询是否该学生是要往哪个表中进行insert 邱慕夏 2015年11月20日16:57:30
        public int InsertExcelRecord(ExcelQuestionRecordEntity studentrecord)
        {
            return exceldal.InsertExcelRecord(studentrecord);
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
            return exceldal.SelectExcelRecord(studentrecord);
        }
        #endregion

        #region 批量根据学生的ID查询是否该学生是要往哪个表中进行insert 邱慕夏 2015年11月20日16:57:30
        public int InsertExcelRecordList(List<ExcelQuestionRecordEntity> studentrecordlist)
        {
            return exceldal.InsertExcelRecordList(studentrecordlist);
        }
        #endregion
    }
}
