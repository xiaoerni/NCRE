using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;
using DAL;
using System.Data.SqlClient;
using System.Data;


namespace BLL
{
    public  class StudentScoreBLL
    {
        private StudentScoreDAL studentscoredal;
        public StudentScoreBLL() {
            studentscoredal = new StudentScoreDAL();
        }



        #region 通过专业的条件，选出本班所有学生的得分信息---周洲--2015年11月14日20:34:18
        
        /// <summary>
        /// 通过专业的条件，选出本班所有学生的得分信息
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectScoreByMajor(StudentInfoEntity studentinfo) {
         
            return studentscoredal.SelectStudentByMajor(studentinfo);

        }
        #endregion

        #region 通过学院的条件，选出本学院所有学生的得分信息---周洲--2015年11月14日20:34:18

        /// <summary>
        /// 通过学院的条件，选出本班所有学生的得分信息
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectScoreByCollege(StudentInfoEntity studentinfo)
        {

            return studentscoredal.SelectScoreByCollege(studentinfo);

        }
        #endregion


        #region 选择所有需要考试学生的专业---周洲--2015年11月16日
        /// <summary>
        /// 选择所有需要考试学生的专业---周洲--2015年11月16日
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectAllMajor() {
            return studentscoredal.SelectAllMajor();
        }
        #endregion

        #region 选择所有需要考试学生的学院---周洲--2015年11月16日
        /// <summary>
        /// 选择所有需要考试学生的专业---周洲--2015年11月16日
        /// </summary>
        /// <returns></returns>
        public DataTable SelectAllCollege()
        {
            return studentscoredal.SelectAllCollege();
        }
        #endregion

        #region 选择所有需要考试学生的学院ID和姓名---周洲--2015年11月16日
        /// <summary>
        /// 选择所有需要考试学生的学院ID和姓名---周洲--2015年11月16日
        /// </summary>
        /// <returns></returns>
        public DataTable SelectAllCollegeInfo()
        {
            return studentscoredal.SelectAllCollegeInfo();
        }
        #endregion

        #region 选择对应学院下拉框的专业---周洲--2015年11月17日
        /// <summary>
        /// 选择对应学院下拉框的专业---周洲--2015年11月17日
        /// </summary>
        /// <param name="studentinfo"></param>
        /// <returns></returns>
        public DataTable SelectMajorByCollegeID(StudentInfoEntity studentinfo)
        {
            return studentscoredal.SelectMajorByCollegeID(studentinfo);
        }
        #endregion

    }
}
