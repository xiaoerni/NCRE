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
    public class WinQuestionEntityBLL
    {
        private WinQuestionEntityDAL winQuestionDal;
        public WinQuestionEntityBLL ()
        {
                //创建一个winQuestionDAL
                winQuestionDal = new DAL.WinQuestionEntityDAL();
         }

        /**
      * qmx
      * **/
        private TBToList<WinQuestionEntity> tBToList = new TBToList<WinQuestionEntity>();

        #region "根据试卷类型，从题库中选出试题的相关信息-韩梦甜-2015-11-20"
        /// <summary>
        /// 根据试卷类型，从题库中选出试题的相关信息
        /// </summary>
        /// <param name="winquestion"></param>
        /// <returns></returns>
        public DataTable LoadWindowsByFlag(WinQuestionEntity winquestion)
        {
            DataTable winQuestionDt = new DataTable();
            winQuestionDt = winQuestionDal.LoadWindowsByFlag(winquestion);

            int num = winQuestionDt.Rows.Count;      //查询到的datatable的行数

            if (num == 0)
            {
                MessageBox.Show("抽题失败，请联系管理员");

            }

            return winQuestionDt;

        }
        #endregion

        #region  根据试题类型获取题库中的试题--周洲--2015年11月21日
        /// <summary>
        ///  #region 根据试题类型获取题库中的试题--周洲--2015年11月21日
        /// </summary>
        /// <param name="studentinfo">放学号的实体</param>
        /// <returns>返回取出的题目</returns>
        public DataTable LoadWindowsQuestion(WinQuestionEntity wininfo)
        {
            return winQuestionDal.LoadWindowsQuestion(wininfo);
        }
        #endregion

        #region"根据学号查询试卷类型-韩梦甜-2015-11-20"
        /// <summary>
        /// 根据学号查询试卷类型
        /// </summary>
        /// <param name="studentinfo">放学号的实体</param>
        /// <returns>试卷类型</returns>
        public DataTable SelectPaperTypeByStudentIDBLL(StudentInfoEntity studentinfo)
        {
            return winQuestionDal.SelectPaperTypeByStudentID(studentinfo);
        }
        #endregion

        #region"返回分值到答题记录表中-韩梦甜-2015-11-20"
        /// <summary>
        /// 返回分值到WinQuestionEntityRecord  表中
        /// </summary>
        /// <param name="studentRecord">正确答案</param>
        /// <returns>返回学号，得分，学生答案，正确答案</returns>
        public int ReturnScore(WinQuestionRecordEntity studentRecord)
        {
            int flag = winQuestionDal.ReturnScore(studentRecord);
            return flag;
        }
        #endregion

        #region 查找所有的word的套卷 WordPaperType() 邱慕夏 2015年11月20日16:57:30
        public DataTable WinPaperType()
        {
            return winQuestionDal.WinPaperType();
        }
        #endregion

        #region 给WordQutionEntity分组 邱慕夏 2015年11月20日16:57:30
        public List<WinQuestionEntity> WinPaperTypeGroupByPaperType(String PaperType)
        {
            return tBToList.ToList(winQuestionDal.WinPaperTypeGroupByPaperType(PaperType));
        }
        #endregion

        #region 根据学生的ID查询去重--邱慕夏  2015年11月23日14:28:57
        /// <summary>
        /// 根据学生的ID查询去重--邱慕夏--邱慕夏
        /// </summary>
        /// <param name="studentinfo">根据CollegeID判断学生在那个表中</param>
        /// <returns>是否添加成功</returns>
        public Boolean SelectWinRecord(WinQuestionRecordEntity studentrecord)
        {
            return winQuestionDal.SelectWinRecord(studentrecord);
        }
        #endregion

        #region 根据学生的ID查询是否该学生是要往哪个表中进行insert 邱慕夏 2015年11月20日16:57:30
        public int InsertWinRecord(WinQuestionRecordEntity studentrecord)
        {
            return winQuestionDal.InsertWinRecord(studentrecord);
        }
        #endregion

        #region 批量根据学生的ID查询是否该学生是要往哪个表中进行insert 邱慕夏 2015年11月20日16:57:30
        public int InsertWinRecordList(List<WinQuestionRecordEntity> studentrecordlist)
        {
            return winQuestionDal.InsertWinRecordList(studentrecordlist);
        }
        #endregion
    }
}
