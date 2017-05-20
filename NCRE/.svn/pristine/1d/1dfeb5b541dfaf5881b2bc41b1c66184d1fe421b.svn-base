using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DAL;
using Model;
using System.Data;
using System.Data.SqlClient;
//using System.Threading.Tasks;
using System.Collections;
using System.Windows.Forms;

namespace BLL
{
    public class IEQuestionEntityBLL
    {
         private IEQuestionEntityDAL ieQuestionDal;
        public IEQuestionEntityBLL ()
        {
                //创建一个winQuestionDAL
                ieQuestionDal = new IEQuestionEntityDAL();
         }

        /**
       * qmx
       * **/
        private TBToList<IEQuestionEntity> tBToList = new TBToList<IEQuestionEntity>();

        #region 根 据题型选择出IE的题干--周洲--2015年11月21日
        /// <summary>
        /// 根据题型选择出IE的题干--周洲--2015年11月21日
        /// </summary>
        /// <param name="iequestion"></param>
        /// <returns></returns>
        public DataTable LoadIEQuestion(IEQuestionEntity ieinfo)
        {
            DataTable ieQuestionDt = new DataTable();
            ieQuestionDt = ieQuestionDal.LoadIEQuestion(ieinfo);

            int num = ieQuestionDt.Rows.Count;      //查询到的datatable的行数

            if (num == 0)
            {
                MessageBox.Show("抽题失败，请联系管理员");

            }


            return ieQuestionDt;

        }
        #endregion

        #region "根据试卷类型，从题库中选出试题的相关信息-韩梦甜-2015-11-20"
        /// <summary>
        /// 根据试卷类型，从题库中选出试题的相关信息
        /// </summary>
        /// <param name="iequestion"></param>
        /// <returns></returns>
        public DataTable LoadIEByFlag(IEQuestionEntity iequestion)
        {
  
            return ieQuestionDal.LoadIEByFlag (iequestion);

        }
        #endregion

        #region "根据学号查询试卷类型-韩梦甜-2015-11-20"
        /// <summary>
        /// 根据学号查询试卷类型
        /// </summary>
        /// <param name="iequestion"></param>
        /// <returns></returns>
        public DataTable SelectPaperTypeByStudentIDBLL(StudentInfoEntity studentinfo)
        {
            return ieQuestionDal.SelectPaperTypeByStudentID(studentinfo);
        }
        #endregion

        #region"返回分值到答题记录表中-韩梦甜-2015-11-20"
        /// <summary>
        /// 返回分值到IEQuestionEntityRecord  表中
        /// </summary>
        /// <param name="studentRecord">正确答案</param>
        /// <returns>返回学号，得分，学生答案，正确答案</returns>
        public int ReturnScore(IEQuestionRecordEntity studentRecord)
        {
            int flag = ieQuestionDal.ReturnScore(studentRecord);
            return flag;
        }
        #endregion

        #region 查找所有的word的套卷 WordPaperType() 邱慕夏 2015年11月20日16:57:30
        public DataTable WinPaperType()
        {
            return ieQuestionDal.IEPaperType();
        }
        #endregion

        #region 给WordQutionEntity分组 邱慕夏 2015年11月20日16:57:30
        public List<IEQuestionEntity> IEPaperTypeGroupByPaperType(String PaperType)
        {
            return tBToList.ToList(ieQuestionDal.IEPaperTypeGroupByPaperType(PaperType));
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
            return ieQuestionDal.SelectIERecord(studentrecord);
        }
        #endregion

        #region 根据学生的ID查询是否该学生是要往哪个表中进行insert 邱慕夏 2015年11月20日16:57:30
        public int InsertIERecord(IEQuestionRecordEntity studentrecord)
        {
            return ieQuestionDal.InsertIERecord(studentrecord);
        }
        #endregion

        #region 批量根据学生的ID查询是否该学生是要往哪个表中进行insert 邱慕夏 2015年11月20日16:57:30
        public int InsertIERecordList(List<IEQuestionRecordEntity> studentrecordlist)
        {
            return ieQuestionDal.InsertIERecordList(studentrecordlist);
        }
        #endregion
    }
}
