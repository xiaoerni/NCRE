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
    /// <summary>
    /// 教师登陆-- 2015年11月14日--周洲
    /// </summary>
    public class UserEntityBLL
    {

        private UserEntityDAL userentitydal = null;
        public UserEntityBLL()
        {
            userentitydal = new UserEntityDAL();
        }


        #region 教师登陆--周洲--2015年11月14日14:38:29
        /// <summary>
        /// 教师登陆--周洲--2015年11月14日14:38:29
        /// </summary>
        /// <param name="userinfo"></param>
        /// <returns></returns>
        public DataTable TeacherLoginByName(UserEntity userinfo)
        {
            return userentitydal.TeacherLoginByName(userinfo);
        }
        #endregion

        #region 查询该考生 是否有资格考试 -赵崇-2015年11月24日 16:46:46
        /// <summary>
        /// 查询该考生 是否有资格考试
        /// </summary>
        /// <param name="pEnStudent">学生</param>
        /// <returns>true表示可以进行考试，false表示 未进行配置 所以不能考试</returns>
        public bool GetIsCanExamByStudent(StudentInfoEntity pEnStudent)
        {
            ExcelEntityDAL excelDal = new ExcelEntityDAL();
            IEQuestionEntityDAL ieDal = new IEQuestionEntityDAL();
            PptQuestionEntityDAL pptDal = new PptQuestionEntityDAL();
            WinQuestionEntityDAL winDal = new WinQuestionEntityDAL();
            WordQuestionEntityDAL wordDal = new WordQuestionEntityDAL();
            SelectQuestionRecordEntityDAL selectDal = new SelectQuestionRecordEntityDAL();

            if (excelDal.GetIsCanExamByStudent(pEnStudent) == false)
            {
                return false;
            }
            if (ieDal.GetIsCanExamByStudent(pEnStudent) == false)
            {
                return false;
            }
            if (pptDal.GetIsCanExamByStudent(pEnStudent) == false)
            {
                return false;
            }
            if (winDal.GetIsCanExamByStudent(pEnStudent) == false)
            {
                return false;
            }
            if (wordDal.GetIsCanExamByStudent(pEnStudent) == false)
            {
                return false;
            }
            if (selectDal.GetIsCanExamByStudent(pEnStudent) == false)
            {
                return false;
            }
            return true;
        } 
        #endregion

    }
}
