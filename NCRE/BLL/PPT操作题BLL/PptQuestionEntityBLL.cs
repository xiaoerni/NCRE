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
    public class PptQuestionEntityBLL

    {

        private PptQuestionEntityDAL pptquestiondal = null;
        public PptQuestionEntityBLL()
        {
            pptquestiondal = new PptQuestionEntityDAL();
        }
        /**
          * qmx
          * **/
        private TBToList<PptQuestionEntity> tBToList = new TBToList<PptQuestionEntity>();


        #region 根据试卷的类型，从题库中选出试题的相关信息--周洲--2015年11月21日
        /// <summary>
        /// 根据试卷的类型，从题库中选出试题的相关信息
        /// </summary>
        /// <param name="studentinfo">放学号的实体</param>
        /// <returns>返回取出的题目</returns>
        public DataTable LoadPptQuestion(PptQuestionEntity pptinfo)
        {
            return pptquestiondal.LoadPptQuestion(pptinfo);
        }

        #endregion


        /// <summary>
        /// 根据试卷的类型和题的类型，从题库中选出试题的相关信息
        /// </summary>
        /// <param name="studentinfo">放学号的实体</param>
        /// <returns>返回取出的题目</returns>
       
        public DataTable LoadPptByFlag(PptQuestionEntity pptinfo)
        {
            return pptquestiondal.LoadPptByFlag(pptinfo);
        }


        /// <summary>
        /// 返回分值到数据库当中
        /// </summary>
        /// <param name="studentrecord">正确答案</param>
        /// <returns>返回学号，得分，学生答案，正确答案</returns>
        /// 
        public int ReturnScore(PptQuestionRecordEntity studentrecord)
        {
            return pptquestiondal.ReturnScord(studentrecord);
        }

        #region 根据学号查询试卷类型 李少然
        public DataTable SelectPaperTypeByStudentIDBLL(StudentInfoEntity studentinfo)
        {
            return pptquestiondal.SelectPaperTypeByStudentID(studentinfo);
        }
        #endregion

        #region 查找所有的Ppt的套卷 PptPaperType() 邱慕夏 2015年11月20日16:57:30
        public DataTable PptPaperType()
        {
            return pptquestiondal.PptPaperType();
        }
        #endregion

        #region 给PptQutionEntity分组 邱慕夏 2015年11月20日16:57:30
        public List<PptQuestionEntity> PptPaperTypeGroupByPaperType(String PaperType)
        {

            return tBToList.ToList(pptquestiondal.PptPaperTypeGroupByPaperType(PaperType));
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
            return pptquestiondal.SelectPptRecord(studentrecord);
        }
        #endregion

        #region 根据学生的ID查询是否该学生是要往哪个表中进行insert 邱慕夏 2015年11月20日16:57:30
        public int InsertPptRecord(PptQuestionRecordEntity studentrecord)
        {
            return pptquestiondal.InsertPptRecord(studentrecord);
        }
        #endregion

        #region 批量根据学生的ID查询是否该学生是要往哪个表中进行insert 邱慕夏 2015年11月20日16:57:30
        public int InsertPptRecordList(List<PptQuestionRecordEntity> studentrecordlist)
        {
            return pptquestiondal.InsertPptRecordList(studentrecordlist);
        }
        #endregion
    }
}
