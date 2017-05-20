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
    public  class WordQuestionEntityBLL
    {
        private WordQuestionEntityDAL wordquestiondal = null;

        public WordQuestionEntityBLL() 
        {
            wordquestiondal = new WordQuestionEntityDAL();
        }

        /**
       * qmx
       * **/
        private TBToList<WordQuestionEntity> tBToList = new TBToList<WordQuestionEntity>();

        #region 利用全局变量，从题库中加载word试题--周洲--2015年11月21日
        /// <summary>
        ///利用全局变量，从题库中加载word试题--周洲--2015年11月21日
        /// </summary>
        /// <param name="studentinfo">放学号的实体</param>
        /// <returns>返回取出的题目</returns>
        public DataTable LoadWordQuestion(WordQuestionEntity wordinfo)
        {
            return wordquestiondal.LoadWordQuestion(wordinfo);
        }
        #endregion

    


        /// <summary>
        /// 根据试卷的类型和题的类型，从题库中选出试题的相关信息
        /// </summary>
        /// <param name="studentinfo">放学号的实体</param>
        /// <returns>返回取出的题目</returns>
        public DataTable LoadWordByFlag(WordQuestionEntity wordinfo)
        {
            return wordquestiondal.LoadWordByFlag (wordinfo);
        }
        


        /// <summary>
        /// 返回分值到数据库当中
        /// </summary>
        /// <param name="studentrecord">正确答案</param>
        /// <returns>返回学号，得分，学生答案，正确答案</returns>
        public int ReturnScore(WordQuestionRecordEntity studentrecord)
        {
             int flag= wordquestiondal.ReturnScore( studentrecord);
             return flag;
        }

        #region 根据学号查询试卷类型 李少然
        public DataTable SelectPaperTypeByStudentIDBLL(StudentInfoEntity studentinfo)
        {
            return wordquestiondal.SelectPaperTypeByStudentID(studentinfo);
        }
        #endregion

        #region 查找所有的word的套卷 WordPaperType() 邱慕夏 2015年11月20日16:57:30
        public DataTable WordPaperType()
        {
            return wordquestiondal.WordPaperType();
        }
        #endregion

        #region 给WordQutionEntity分组 邱慕夏 2015年11月20日16:57:30
        public List<WordQuestionEntity> WordPaperTypeGroupByPaperType(String PaperType)
        {

            return tBToList.ToList(wordquestiondal.WordPaperTypeGroupByPaperType(PaperType));
        }
        #endregion

        #region 根据学生的ID查询是否该学生是要往哪个表中进行insert 邱慕夏 2015年11月20日16:57:30
        public int InsertWordRecord(WordQuestionRecordEntity studentrecord)
        {
            return wordquestiondal.InsertWordRecord(studentrecord);
        }
        #endregion

        #region 根据学生的ID查询是否该学生是要往哪个表中进行insert--邱慕夏  2015年11月23日14:28:57
        /// <summary>
        /// 根据学生的ID查询该学生是否已经存在WordRecordQuestionEntity中--邱慕夏
        /// </summary>
        /// <param name="studentinfo">根据CollegeID判断学生在那个表中</param>
        /// <returns>是否添加成功</returns>
        public Boolean SelectWordRecord(WordQuestionRecordEntity studentrecord)
        {
            return wordquestiondal.SelectWordRecord(studentrecord);
        }
        #endregion


        #region 批量根据学生的ID查询是否该学生是要往哪个表中进行insert 邱慕夏 2015年11月20日16:57:30
        public int InsertWordRecordList(List<WordQuestionRecordEntity> studentrecordlist)
        {
            return wordquestiondal.InsertWordRecordList(studentrecordlist);
        }
        #endregion


    }
}
