using DAL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;

namespace BLL
{
    public class SelectQuestionBLL
    {
        //创建一个D层的实例
        SelectQuestionRecordEntityDAL sqDal = new SelectQuestionRecordEntityDAL();

        #region 查询学生的答题记录
        /// <summary>
        /// 查询学生的答题记录
        /// </summary>
        /// <param name="pEnStudent"></param>
        /// <returns></returns>
        public List<SelectQuestionRecordEntity> GetLstSelectQuestionRecordByStudentIdAndCollegeId(StudentInfoEntity pEnStudent)
        {
            List<SelectQuestionRecordEntity> listSqRecord = new List<SelectQuestionRecordEntity>();
            listSqRecord = sqDal.GetLstSelectQuestionRecordByStudentIdAndCollegeId(pEnStudent);
            if (listSqRecord != null && listSqRecord.Count > 0)
            {
                return listSqRecord;
            }
            else
            {
                return null;
            }
        }
        #endregion

        #region 更新选择题答题记录UpdateSelectQuestionRecordByStudentInfo
        /// <summary>
        /// 更新选择题答题记录
        /// </summary>
        /// <param name="pEnStudentInfo">学生实体</param>
        /// <param name="pEnSelectRecord">选择题答题记录实体</param>
        /// <returns></returns>
        public int UpdateSelectQuestionRecordByStudentInfo(StudentInfoEntity pEnStudentInfo, SelectQuestionRecordEntity pEnSelectRecord,string rightAnswer)
        {
            int flag = sqDal.UpdateSelectQuestionRecordByStudentInfo(pEnStudentInfo, pEnSelectRecord,rightAnswer);
            if (flag == 0)
            {
                return 0;
            }
            else
            {
                return 1;
            }
        } 
        #endregion

        #region 查询该考生 是否有资格考试 -赵崇-2015年11月24日
        /// <summary>
        /// 查询该考生 是否有资格考试
        /// </summary>
        /// <param name="pEnStudent">学生</param>
        /// <returns>true表示可以进行考试，false表示 未进行配置 所以不能考试</returns>
        public bool GetIsCanExamByStudent(StudentInfoEntity pEnStudent)
        {
            return sqDal.GetIsCanExamByStudent(pEnStudent);
        } 
        #endregion
    }
}
